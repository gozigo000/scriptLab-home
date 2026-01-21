Option Explicit

' ============================================================
' 모듈: outlineCache
' 역할: 문서 개요(탐색창 제목) 캐시 빌더
' Public 함수/서브:
'   - GetCurrentHeadingTitle: 현재 위치의 제목 문자열 반환
'   - GetCurrentHeadingLevel: 현재 위치의 제목 레벨 반환
'   - GetCurrentHeadingRange: 현재 위치의 제목 구간(Range) 반환
'   - DeleteOutlineCache: 캐시 무효화
'   - (디버깅용) a_ShowOutlineHeadingInfo: 현재 위치의 제목 정보 알림
'   - (디버깅용) a_DeleteCurrentDocumentOutlineCache: 현재 문서 캐시 무효화
' ============================================================
'
' Lazy 캐시:
' - 문서 전체를 한 번에 스캔하지 않습니다(초반 빌드 비용 제거).
' - 커서가 "다른 아웃라인 영역(제목 단위)"로 이동했을 때만,
'   현재 커서가 속한 영역의 제목/경계를 계산해서 캐시합니다.
'
' 캐시 범위 정의(섹션):
' - headingStart <= pos < nextHeadingStart
' - headingStart: "현재 위치에서 위로 가장 가까운 제목" 문단의 Start 위치
' - nextHeadingStart: "현재 위치에서 아래로 가장 가까운 다음 제목" 문단의 Start 위치
' (즉, 레벨과 무관하게 다음 제목(Heading1/2/3...)이 나오면 섹션이 바뀐 것으로 간주)
'
' 제목 판정 기준:
' - Paragraph.OutlineLevel <> wdOutlineLevelBodyText
'
' 사용 예:
' - title = GetCurrentHeadingTitle(Selection.Range, 140)
' - Call DeleteOutlineCache() ' 필요 시 강제 갱신/초기화

' ===== Lazy 캐시 저장소(모듈 레벨) =====
Private gDocKey As String
Private gFingerprint As String

' ===== 최근 N개 문서 RAM 캐시(LRU) =====
' - 문서를 왔다갔다 할 때 CustomXMLParts 파싱을 반복하지 않도록,
'   최근 문서들의 캐시를 메모리에 유지합니다.
Private Const MAX_DOC_RAM_CACHE As Long = 5
Private gRamDocCount As Long
Private gRamDocKey() As String
Private gRamDocFp() As String
Private gRamDocLastAccess() As Date
Private gRamSecCount() As Long
Private gRamSecStart() As Variant   ' Variant holding Long() or Empty
Private gRamSecEnd() As Variant     ' Variant holding Long() or Empty
Private gRamSecLevel() As Variant   ' Variant holding Long() or Empty
Private gRamSecTitle() As Variant   ' Variant holding String() or Empty
Private gRamLastIdx() As Long

' 여러 섹션(구간) 캐시를 유지해서 섹션을 오가도 빠르게 찾도록 함
Private Const MAX_SECTION_CACHE As Long = 256
Private gSecCount As Long
Private gSecStart() As Long        ' 섹션 시작(제목 문단 Start) - 오름차순 유지
Private gSecEnd() As Long          ' 섹션 끝(포함) = nextHeadingStart-1
Private gSecLevel() As Long
Private gSecTitle() As String
Private gLastIdx As Long           ' 직전 히트 인덱스(로컬리티 최적화)

' ===== 문서 내 저장(CustomXMLParts) 설정 =====
Private Const OUTLINE_XML_NS As String = "urn:scriptlab:outlinecache:v1"
Private Const OUTLINE_XML_ROOT_LOCAL As String = "outlineCache"
Private Const OUTLINE_XML_VERSION As String = "1"

' ===== Public API =====

' 현재 커서가 속한 문단에서 위쪽으로 가장 가까운
' "제목(탐색창에 표시되는 개요 수준)"을 찾아 반환
' - rng: 현재 커서가 속한 문단의 Range
' - maxLen: 반환 문자열의 최대 길이
' - 반환값: "제목" (또는 "1.2.3 제목" 처럼 목록번호가 있으면 포함)
' 하위호환용 이름: 내부적으로 lazy 캐시 사용
Public Function GetCurrentHeadingTitle( _
    ByVal rng As Range, _
    Optional ByVal maxLen As Long = 140 _
) As String
    GetCurrentHeadingTitle = GetCurrentHeadingTitleLazy(rng, maxLen)
End Function

' 현재 커서가 속한 문단의 "제목 레벨(OutlineLevel)"을 반환
' - rng: 현재 커서가 속한 문단의 Range
' - 반환값: WdOutlineLevel 값을 Long으로 반환(없으면 0)
Public Function GetCurrentHeadingLevel( _
    ByVal rng As Range _
) As Long
    GetCurrentHeadingLevel = GetCurrentHeadingLevelLazy(rng)
End Function

' 현재 위치 기준:
' - 위로 가장 가까운 제목(헤딩) 문단 시작 ~
' - 아래로 가장 가까운 다음 제목(헤딩) 문단 시작 직전(End는 다음 제목 Start로 설정)
' 을 Range로 반환합니다.
' - 제목을 찾지 못하면 Nothing 반환
Public Function GetCurrentHeadingRange( _
    ByVal rng As Range _
) As Range
    Set GetCurrentHeadingRange = GetCurrentHeadingRangeLazy(rng)
End Function

' (MARK) (디버깅용) 현재 위치의 "제목(개요)" 정보 알림
' ----------------------
' 현재 커서(Selection) 기준으로 다음 정보를 알림창으로 표시합니다.
' - 제목 문자열(탐색창에 표시되는 개요 제목)
' - 제목 레벨(OutlineLevel)
' - 제목 구간(시작/끝 Range)
'
' 사용:
' - 매크로: a_ShowOutlineHeadingInfo 실행
Public Sub a_ShowOutlineHeadingInfo()
    On Error GoTo ErrorHandler
    
    Dim selRng As Range
    Set selRng = Selection.Range
    If selRng Is Nothing Then
        VBA.MsgBox "Selection.Range를 가져올 수 없습니다.", vbInformation, "현재 제목 정보"
        Exit Sub
    End If
    
    Dim headingTitle As String
    headingTitle = GetCurrentHeadingTitle(selRng, 200)
    
    Dim headingLevel As Long
    headingLevel = GetCurrentHeadingLevel(selRng)
    
    Dim headingRng As Range
    Set headingRng = GetCurrentHeadingRange(selRng)
    
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim rangeLen As Long
    
    If headingRng Is Nothing Then
        rangeStart = 0
        rangeEnd = 0
        rangeLen = 0
    Else
        rangeStart = headingRng.Start
        rangeEnd = headingRng.End
        rangeLen = rangeEnd - rangeStart
    End If
    
    Dim msg As String
    msg = ""
    msg = msg & "제목: " & IIf(headingTitle = "", "(없음)", headingTitle) & vbCrLf
    msg = msg & "레벨: " & CStr(headingLevel) & vbCrLf
    msg = msg & "구간: [" & CStr(rangeStart) & ", " & CStr(rangeEnd) & ")" & _
        " (len=" & CStr(rangeLen) & ")" & vbCrLf
    
    VBA.MsgBox msg, vbInformation, "현재 제목 정보"
    Exit Sub
    
ErrorHandler:
    VBA.MsgBox "오류: " & Err.Description, vbCritical, "현재 제목 정보"
End Sub

' Lazy 방식: rng 위치에서 요청한 경우에만 제목/경계를 갱신합니다.
Private Function GetCurrentHeadingTitleLazy( _
    ByVal rng As Range, _
    Optional ByVal maxLen As Long = 140 _
) As String
    On Error GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document
    
    Dim key As String
    key = MakeDocKey(doc)
    
    Dim fp As String
    fp = MakeFingerprint(doc)
    
    Dim pos As Long
    pos = rng.Start

    ' 문서 전환 시: 최근 N개 RAM 캐시에서 우선 복구(없으면 CustomXMLParts에서 로드)
    EnsureActiveDocumentCache doc, key, fp
    
    ' 1) 최근 히트 섹션이면 O(1)
    If gLastIdx > 0 And gLastIdx <= gSecCount Then
        If pos >= gSecStart(gLastIdx) _
            And pos <= gSecEnd(gLastIdx) Then
            GetCurrentHeadingTitleLazy = FormatHeadingTitle( _
                gSecTitle(gLastIdx), _
                maxLen _
            )
            TouchRamCacheEntry gDocKey, gFingerprint
            Exit Function
        End If
    End If
    
    ' 2) 캐시된 섹션 구간에서 이진탐색
    Dim idx As Long
    idx = FindCachedSectionIndex(pos)
    If idx > 0 Then
        gLastIdx = idx
        GetCurrentHeadingTitleLazy = FormatHeadingTitle( _
            gSecTitle(idx), _
            maxLen _
        )
        TouchRamCacheEntry gDocKey, gFingerprint
        Exit Function
    End If
    
    ' 3) 캐시 미스: 현재 위치 섹션을 계산하고 캐시에 추가
    Dim headingStart As Long
    Dim nextHeadingStart As Long
    Dim headingLevel As Long
    Dim headingTitle As String
    ComputeSectionForRange doc, rng, headingStart, _
        nextHeadingStart, headingLevel, headingTitle
    
    If headingStart > 0 Then
        idx = AddOrUpdateSectionCache( _
            headingStart, nextHeadingStart, _
            headingLevel, headingTitle _
        )
        gLastIdx = idx
        GetCurrentHeadingTitleLazy = FormatHeadingTitle( _
            headingTitle, maxLen _
        )
        ' 캐시가 늘어났으면 문서에 저장
        ' (초기 전체 스캔이 없어서 빈도는 낮은 편)
        Call SaveCacheToDocument(doc)
        ' RAM 캐시도 최신 상태로 반영
        UpsertRamCacheFromCurrent
    Else
        GetCurrentHeadingTitleLazy = ""
    End If
    Exit Function
    
SafeExit:
    GetCurrentHeadingTitleLazy = ""
End Function

' Lazy 방식: rng 위치에서 요청한 경우에만 제목/경계를 갱신합니다.
Private Function GetCurrentHeadingLevelLazy( _
    ByVal rng As Range _
) As Long
    On Error GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document
    
    Dim key As String
    key = MakeDocKey(doc)
    
    Dim fp As String
    fp = MakeFingerprint(doc)
    
    Dim pos As Long
    pos = rng.Start

    ' 문서 전환 시: 최근 N개 RAM 캐시에서 우선 복구(없으면 CustomXMLParts에서 로드)
    EnsureActiveDocumentCache doc, key, fp
    
    ' 1) 최근 히트 섹션이면 O(1)
    If gLastIdx > 0 And gLastIdx <= gSecCount Then
        If pos >= gSecStart(gLastIdx) _
            And pos <= gSecEnd(gLastIdx) Then
            GetCurrentHeadingLevelLazy = gSecLevel(gLastIdx)
            TouchRamCacheEntry gDocKey, gFingerprint
            Exit Function
        End If
    End If
    
    ' 2) 캐시된 섹션 구간에서 이진탐색
    Dim idx As Long
    idx = FindCachedSectionIndex(pos)
    If idx > 0 Then
        gLastIdx = idx
        GetCurrentHeadingLevelLazy = gSecLevel(idx)
        TouchRamCacheEntry gDocKey, gFingerprint
        Exit Function
    End If
    
    ' 3) 캐시 미스: 현재 위치 섹션을 계산하고 캐시에 추가
    Dim headingStart As Long
    Dim nextHeadingStart As Long
    Dim headingLevel As Long
    Dim headingTitle As String
    ComputeSectionForRange doc, rng, headingStart, _
        nextHeadingStart, headingLevel, headingTitle
    
    If headingStart > 0 Then
        idx = AddOrUpdateSectionCache( _
            headingStart, nextHeadingStart, _
            headingLevel, headingTitle _
        )
        gLastIdx = idx
        GetCurrentHeadingLevelLazy = headingLevel
        ' 캐시가 늘어났으면 문서에 저장
        ' (초기 전체 스캔이 없어서 빈도는 낮은 편)
        Call SaveCacheToDocument(doc)
        ' RAM 캐시도 최신 상태로 반영
        UpsertRamCacheFromCurrent
    Else
        GetCurrentHeadingLevelLazy = 0
    End If
    
    Exit Function
    
SafeExit:
    GetCurrentHeadingLevelLazy = 0
End Function

' Lazy 방식: rng 위치에서 요청한 경우에만 제목/경계를 갱신합니다.
Private Function GetCurrentHeadingRangeLazy( _
    ByVal rng As Range _
) As Range
    On Error GoTo SafeExit
    
    Dim doc As Document
    Set doc = rng.Document
    
    Dim key As String
    key = MakeDocKey(doc)
    
    Dim fp As String
    fp = MakeFingerprint(doc)
    
    Dim pos As Long
    pos = rng.Start
    
    ' 문서 전환 시: 최근 N개 RAM 캐시에서 우선 복구(없으면 CustomXMLParts에서 로드)
    EnsureActiveDocumentCache doc, key, fp
    
    Dim headingStart As Long
    headingStart = 0
    
    ' 1) 최근 히트 섹션이면 O(1) (섹션 시작=현재 헤딩 시작)
    If gLastIdx > 0 And gLastIdx <= gSecCount Then
        If pos >= gSecStart(gLastIdx) _
            And pos <= gSecEnd(gLastIdx) Then
            headingStart = gSecStart(gLastIdx)
        End If
    End If
    
    ' 2) 캐시된 섹션 구간에서 이진탐색
    If headingStart = 0 Then
        Dim idx As Long
        idx = FindCachedSectionIndex(pos)
        If idx > 0 Then
            gLastIdx = idx
            headingStart = gSecStart(idx)
        End If
    End If
    
    ' 3) 캐시 미스: 섹션을 계산하고 캐시에 추가(이후 호출 최적화)
    If headingStart = 0 Then
        Dim nextHeadingStart As Long
        Dim headingLevel As Long
        Dim headingTitle As String
        ComputeSectionForRange doc, rng, headingStart, _
            nextHeadingStart, headingLevel, headingTitle
        
        If headingStart > 0 Then
            Dim i2 As Long
            i2 = AddOrUpdateSectionCache( _
                headingStart, nextHeadingStart, _
                headingLevel, headingTitle _
            )
            gLastIdx = i2
            Call SaveCacheToDocument(doc)
            UpsertRamCacheFromCurrent
        End If
    End If
    
    If headingStart = 0 Then GoTo SafeExit
    
    ' 4) 아래로 가장 가까운 "다음 제목" 찾기 (레벨 무관, 가장 먼저 만나는 제목)
    Dim nextHeadingStartComputed As Long
    nextHeadingStartComputed = 0
    
    Dim q As Paragraph
    Set q = rng.Paragraphs(1).Next
    Do While Not q Is Nothing
        Dim lvl As WdOutlineLevel
        lvl = q.OutlineLevel
        If lvl <> wdOutlineLevelBodyText Then
            nextHeadingStartComputed = q.Range.Start
            Exit Do
        End If
        On Error Resume Next
        Set q = q.Next
        On Error GoTo SafeExit
    Loop
    
    If nextHeadingStartComputed = 0 Then nextHeadingStartComputed = doc.Content.End
    If nextHeadingStartComputed < headingStart Then nextHeadingStartComputed = headingStart
    
    Set GetCurrentHeadingRangeLazy = doc.Range(headingStart, nextHeadingStartComputed)
    TouchRamCacheEntry gDocKey, gFingerprint
    Exit Function
    
SafeExit:
    Set GetCurrentHeadingRangeLazy = Nothing
End Function

' (MARK) (디버깅용) 현재 문서 캐시 무효화
Public Sub a_DeleteCurrentDocumentOutlineCache()
    Call DeleteOutlineCache(ActiveDocument)
End Sub

' 캐시를 강제로 무효화합니다. (다음 호출 시 재빌드됨)
Public Sub DeleteOutlineCache(Optional ByVal doc As Document)
    On Error GoTo SafeExit
    
    ' 기본은 현재(활성) 문서
    Dim target As Document
    If doc Is Nothing Then
        Set target = ActiveDocument
    Else
        Set target = doc
    End If
    
    Dim key As String
    key = MakeDocKey(target)
    
    Dim fp As String
    fp = MakeFingerprint(target)
    
    ' 1) 문서 내 저장된 캐시(CustomXMLParts) 삭제
    DeleteExistingOutlineParts target
    
    ' 2) RAM(LRU) 캐시에서 해당 문서만 제거
    RemoveRamCacheEntry key, fp
    
    ' 3) 현재 활성 캐시가 같은 문서면 즉시 무효화
    If gDocKey = key And gFingerprint = fp Then
        ResetSectionCache "", ""
        gLastIdx = 0
    End If
    
    Exit Sub
    
SafeExit:
    ' 실패하더라도 최소한 메모리 캐시만 초기화(문서 보호/권한/환경 차이 등)
    ResetSectionCache "", ""
    ResetRamCache
End Sub


' ===== 내부 구현 =====

Private Sub ResetSectionCache(ByVal key As String, ByVal fp As String)
    gDocKey = key
    gFingerprint = fp
    gSecCount = 0
    gLastIdx = 0
    Erase gSecStart
    Erase gSecEnd
    Erase gSecLevel
    Erase gSecTitle
End Sub

' -------------------------
' RAM 캐시 (최근 N개 문서)
' -------------------------

Private Sub ResetRamCache()
    gRamDocCount = 0
    Erase gRamDocKey
    Erase gRamDocFp
    Erase gRamDocLastAccess
    Erase gRamSecCount
    Erase gRamSecStart
    Erase gRamSecEnd
    Erase gRamSecLevel
    Erase gRamSecTitle
    Erase gRamLastIdx
End Sub

' RAM 캐시에서 특정 문서(key/fp)만 제거(슬롯은 비워두고 LRU로 재사용되도록 함)
Private Sub RemoveRamCacheEntry(ByVal key As String, ByVal fp As String)
    On Error GoTo SafeExit
    If gRamDocCount <= 0 Then Exit Sub
    If Not IsArrayAllocated(gRamDocKey) Then Exit Sub
    
    Dim idx As Long
    idx = FindRamCacheIndex(key, fp)
    If idx <= 0 Then Exit Sub
    
    gRamDocKey(idx) = ""
    gRamDocFp(idx) = ""
    gRamDocLastAccess(idx) = #1/1/1900#
    gRamSecCount(idx) = 0
    gRamSecStart(idx) = Empty
    gRamSecEnd(idx) = Empty
    gRamSecLevel(idx) = Empty
    gRamSecTitle(idx) = Empty
    gRamLastIdx(idx) = 0
    
    Exit Sub
SafeExit:
End Sub

' 현재 문서(key/fp)의 캐시가 활성화되도록 보장:
' - RAM에 있으면 즉시 복구
' - 없으면 CustomXMLParts에서 로드(유효하면) 후 RAM에 저장
Private Sub EnsureActiveDocumentCache( _
    ByVal doc As Document, _
    ByVal key As String, _
    ByVal fp As String _
)
    On Error GoTo SafeExit
    
    ' 이미 같은 문서/상태면 접근 시간만 갱신
    If gDocKey = key And gFingerprint = fp Then
        TouchRamCacheEntry key, fp
        Exit Sub
    End If
    
    ' 현재 활성 캐시를 RAM에 저장
    UpsertRamCacheFromCurrent
    
    ' RAM에서 복구 시도
    If LoadFromRamCache(key, fp) Then
        Exit Sub
    End If
    
    ' 없으면 새로 초기화하고, 문서 내 저장분(CustomXMLParts) 로드 시도
    ResetSectionCache key, fp
    Call TryLoadCacheFromDocument(doc, key, fp)
    
    ' 로드 결과(있든 없든)를 RAM에 저장
    UpsertRamCacheFromCurrent
    Exit Sub
    
SafeExit:
    ' 실패하면 최소한 현재 문서키는 맞춰둠
    ResetSectionCache key, fp
End Sub

' 현재 모듈의 활성 캐시(gDocKey/gFingerprint/gSec*)를 RAM 슬롯에 저장/갱신
Private Sub UpsertRamCacheFromCurrent()
    On Error GoTo SafeExit
    
    If gDocKey = "" Or gFingerprint = "" Then Exit Sub
    
    Dim idx As Long
    idx = FindRamCacheIndex(gDocKey, gFingerprint)
    
    If idx = 0 Then
        idx = AllocateRamSlot()
        gRamDocKey(idx) = gDocKey
        gRamDocFp(idx) = gFingerprint
    End If
    
    gRamDocLastAccess(idx) = Now
    gRamSecCount(idx) = gSecCount
    gRamLastIdx(idx) = gLastIdx
    
    If gSecCount <= 0 Then
        gRamSecStart(idx) = Empty
        gRamSecEnd(idx) = Empty
        gRamSecLevel(idx) = Empty
        gRamSecTitle(idx) = Empty
    Else
        gRamSecStart(idx) = gSecStart
        gRamSecEnd(idx) = gSecEnd
        gRamSecLevel(idx) = gSecLevel
        gRamSecTitle(idx) = gSecTitle
    End If
    
    Exit Sub
    
SafeExit:
End Sub

' RAM 슬롯에서 현재 캐시로 복구
Private Function LoadFromRamCache( _
    ByVal key As String, _
    ByVal fp As String _
) As Boolean
    On Error GoTo SafeExit
    
    Dim idx As Long
    idx = FindRamCacheIndex(key, fp)
    If idx = 0 Then GoTo SafeExit
    
    gDocKey = key
    gFingerprint = fp
    gSecCount = gRamSecCount(idx)
    gLastIdx = gRamLastIdx(idx)
    
    If gSecCount <= 0 Or IsEmpty(gRamSecStart(idx)) Then
        gSecCount = 0
        gLastIdx = 0
        Erase gSecStart
        Erase gSecEnd
        Erase gSecLevel
        Erase gSecTitle
    Else
        gSecStart = gRamSecStart(idx)
        gSecEnd = gRamSecEnd(idx)
        gSecLevel = gRamSecLevel(idx)
        gSecTitle = gRamSecTitle(idx)
    End If
    
    gRamDocLastAccess(idx) = Now
    LoadFromRamCache = True
    Exit Function
    
SafeExit:
    LoadFromRamCache = False
End Function

' LRU 접근시간 갱신
Private Sub TouchRamCacheEntry(ByVal key As String, ByVal fp As String)
    On Error Resume Next
    Dim idx As Long
    idx = FindRamCacheIndex(key, fp)
    If idx > 0 Then gRamDocLastAccess(idx) = Now
End Sub

Private Function FindRamCacheIndex( _
    ByVal key As String, _
    ByVal fp As String _
) As Long
    Dim i As Long
    For i = 1 To gRamDocCount
        If gRamDocKey(i) = key And gRamDocFp(i) = fp Then
            FindRamCacheIndex = i
            Exit Function
        End If
    Next i
    FindRamCacheIndex = 0
End Function

Private Function AllocateRamSlot() As Long
    ' 빈 슬롯이 있으면 사용, 없으면 LRU 슬롯을 덮어씀
    Dim idx As Long
    
    If gRamDocCount < MAX_DOC_RAM_CACHE Then
        gRamDocCount = gRamDocCount + 1
        EnsureRamArraysSized gRamDocCount
        AllocateRamSlot = gRamDocCount
        Exit Function
    End If
    
    idx = FindLruRamIndex()
    If idx <= 0 Then idx = 1
    
    ' 덮어쓰기 전에 내용 정리
    gRamDocKey(idx) = ""
    gRamDocFp(idx) = ""
    gRamDocLastAccess(idx) = #1/1/1900#
    gRamSecCount(idx) = 0
    gRamSecStart(idx) = Empty
    gRamSecEnd(idx) = Empty
    gRamSecLevel(idx) = Empty
    gRamSecTitle(idx) = Empty
    gRamLastIdx(idx) = 0
    
    AllocateRamSlot = idx
End Function

Private Sub EnsureRamArraysSized(ByVal n As Long)
    If n <= 0 Then Exit Sub
    If gRamDocCount = 1 And (Not IsArrayAllocated(gRamDocKey)) Then
        ReDim gRamDocKey(1 To n)
        ReDim gRamDocFp(1 To n)
        ReDim gRamDocLastAccess(1 To n)
        ReDim gRamSecCount(1 To n)
        ReDim gRamSecStart(1 To n)
        ReDim gRamSecEnd(1 To n)
        ReDim gRamSecLevel(1 To n)
        ReDim gRamSecTitle(1 To n)
        ReDim gRamLastIdx(1 To n)
    Else
        ReDim Preserve gRamDocKey(1 To n)
        ReDim Preserve gRamDocFp(1 To n)
        ReDim Preserve gRamDocLastAccess(1 To n)
        ReDim Preserve gRamSecCount(1 To n)
        ReDim Preserve gRamSecStart(1 To n)
        ReDim Preserve gRamSecEnd(1 To n)
        ReDim Preserve gRamSecLevel(1 To n)
        ReDim Preserve gRamSecTitle(1 To n)
        ReDim Preserve gRamLastIdx(1 To n)
    End If
End Sub

Private Function FindLruRamIndex() As Long
    Dim i As Long
    Dim bestIdx As Long
    Dim bestTime As Date
    
    bestIdx = 1
    bestTime = gRamDocLastAccess(1)
    
    For i = 2 To gRamDocCount
        If gRamDocLastAccess(i) < bestTime Then
            bestTime = gRamDocLastAccess(i)
            bestIdx = i
        End If
    Next i
    
    FindLruRamIndex = bestIdx
End Function

Private Function IsArrayAllocated(ByVal v As Variant) As Boolean
    On Error GoTo NotAllocated
    If IsArray(v) Then
        Dim lb As Long
        lb = LBound(v)
        IsArrayAllocated = True
        Exit Function
    End If
NotAllocated:
    IsArrayAllocated = False
End Function

' 현재 위치(rng)가 속한 섹션(제목/경계)을 계산 (캐시에 "저장"하지는 않음)
Private Sub ComputeSectionForRange( _
    ByVal doc As Document, _
    ByVal rng As Range, _
    ByRef headingStart As Long, _
    ByRef nextHeadingStart As Long, _
    ByRef headingLevel As Long, _
    ByRef headingTitle As String _
)
    On Error GoTo SafeExit
    
    headingStart = 0
    nextHeadingStart = doc.Content.End + 1
    headingLevel = 0
    headingTitle = ""
    
    ' 1) 위로 올라가며 가장 가까운 "제목 문단" 찾기
    Dim p As Paragraph
    Dim lvl As WdOutlineLevel
    
    Set p = rng.Paragraphs(1)
    Do While Not p Is Nothing
        lvl = p.OutlineLevel
        If lvl <> wdOutlineLevelBodyText Then
            headingStart = p.Range.Start
            headingLevel = CLng(lvl)
            Dim t As String
            t = NormalizeInlineText(p.Range.Text)
            
            Dim listNo As String
            listNo = GetParagraphListNumber(p)
            
            If listNo <> "" Then
                headingTitle = listNo & " " & t
            Else
                headingTitle = t
            End If
            Exit Do
        End If
        On Error Resume Next
        Set p = p.Previous
        On Error GoTo SafeExit
    Loop
    
    If headingStart = 0 Then Exit Sub
    
    ' 2) 아래로 내려가며 섹션 경계(다음 제목: 레벨 무관) 찾기
    Dim q As Paragraph
    Set q = doc.Range(headingStart, headingStart).Paragraphs(1).Next
    
    Do While Not q Is Nothing
        lvl = q.OutlineLevel
        If lvl <> wdOutlineLevelBodyText Then
            nextHeadingStart = q.Range.Start
            Exit Do
        End If
        On Error Resume Next
        Set q = q.Next
        On Error GoTo SafeExit
    Loop
    
    Exit Sub
    
SafeExit:
    headingStart = 0
    nextHeadingStart = 0
    headingLevel = 0
    headingTitle = ""
End Sub

' pos가 속한 캐시 섹션 인덱스 반환(없으면 0)
Private Function FindCachedSectionIndex(ByVal pos As Long) As Long
    Dim lo As Long, hi As Long, mid As Long, cand As Long
    If gSecCount <= 0 Then
        FindCachedSectionIndex = 0
        Exit Function
    End If
    
    ' 섹션 시작 배열에서 pos 이하인 마지막 인덱스를 찾고,
    ' 그 구간에 포함되는지 확인
    lo = 1
    hi = gSecCount
    cand = 0
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If gSecStart(mid) <= pos Then
            cand = mid
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop
    
    If cand > 0 Then
        If pos <= gSecEnd(cand) Then
            FindCachedSectionIndex = cand
        Else
            FindCachedSectionIndex = 0
        End If
    Else
        FindCachedSectionIndex = 0
    End If
End Function

' 섹션(headingStart~nextHeadingStart-1)을 캐시에
' 삽입/업데이트하고 인덱스를 반환
Private Function AddOrUpdateSectionCache( _
    ByVal headingStart As Long, _
    ByVal nextHeadingStart As Long, _
    ByVal headingLevel As Long, _
    ByVal headingTitle As String _
) As Long
    Dim secStart As Long, secEnd As Long
    secStart = headingStart
    secEnd = nextHeadingStart - 1
    If secEnd < secStart Then secEnd = secStart
    
    ' 이미 같은 start가 있으면 갱신
    Dim i As Long
    For i = 1 To gSecCount
        If gSecStart(i) = secStart Then
            gSecEnd(i) = secEnd
            gSecLevel(i) = headingLevel
            gSecTitle(i) = headingTitle
            AddOrUpdateSectionCache = i
            Exit Function
        End If
    Next i
    
    ' 캐시 상한(간단한 정책): 너무 커지면 앞쪽(가장 오래된 start) 일부를 잘라냄
    If gSecCount >= MAX_SECTION_CACHE Then
        TrimOldestSections MAX_SECTION_CACHE \ 4 ' 1/4 제거
    End If
    
    ' 오름차순 삽입 위치 찾기
    Dim ins As Long
    ins = FindInsertPos(secStart)
    
    ' 배열 확장
    gSecCount = gSecCount + 1
    ReDim Preserve gSecStart(1 To gSecCount)
    ReDim Preserve gSecEnd(1 To gSecCount)
    ReDim Preserve gSecLevel(1 To gSecCount)
    ReDim Preserve gSecTitle(1 To gSecCount)
    
    ' 뒤로 밀기
    For i = gSecCount To ins + 1 Step -1
        gSecStart(i) = gSecStart(i - 1)
        gSecEnd(i) = gSecEnd(i - 1)
        gSecLevel(i) = gSecLevel(i - 1)
        gSecTitle(i) = gSecTitle(i - 1)
    Next i
    
    ' 삽입
    gSecStart(ins) = secStart
    gSecEnd(ins) = secEnd
    gSecLevel(ins) = headingLevel
    gSecTitle(ins) = headingTitle
    
    ' 인접/중복 구간 정리(간단히 좌우 1칸만 체크)
    MergeNeighbors ins
    
    AddOrUpdateSectionCache = ins
End Function

Private Function FindInsertPos(ByVal secStart As Long) As Long
    ' gSecStart 오름차순에서 삽입 위치(1..gSecCount+1)
    Dim lo As Long, hi As Long, mid As Long, ans As Long
    lo = 1
    hi = gSecCount
    ans = gSecCount + 1
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If gSecStart(mid) >= secStart Then
            ans = mid
            hi = mid - 1
        Else
            lo = mid + 1
        End If
    Loop
    FindInsertPos = ans
End Function

Private Sub MergeNeighbors(ByVal idx As Long)
    ' 같은 섹션을 중복 저장하거나 겹치는 경우를 최소화 (간단히 인접 1칸 병합)
    If gSecCount <= 1 Then Exit Sub
    
    ' 왼쪽과 겹치면 왼쪽을 제거하고 idx를 하나 당김(새 섹션을 유지)
    If idx > 1 Then
        If gSecEnd(idx - 1) >= gSecStart(idx) Then
            ' 겹침: 더 넓은 end를 유지,
            ' title/level은 "최근 삽입(idx)" 것을 유지
            If gSecEnd(idx - 1) > gSecEnd(idx) Then
                gSecEnd(idx) = gSecEnd(idx - 1)
            End If
            DeleteSectionAt idx - 1
            idx = idx - 1
        End If
    End If
    
    ' 오른쪽과 겹치면 오른쪽 제거
    If idx < gSecCount Then
        If gSecEnd(idx) >= gSecStart(idx + 1) Then
            If gSecEnd(idx + 1) > gSecEnd(idx) Then
                gSecEnd(idx) = gSecEnd(idx + 1)
            End If
            DeleteSectionAt idx + 1
        End If
    End If
End Sub

Private Sub DeleteSectionAt(ByVal idx As Long)
    Dim i As Long
    If idx < 1 Or idx > gSecCount Then Exit Sub
    For i = idx To gSecCount - 1
        gSecStart(i) = gSecStart(i + 1)
        gSecEnd(i) = gSecEnd(i + 1)
        gSecLevel(i) = gSecLevel(i + 1)
        gSecTitle(i) = gSecTitle(i + 1)
    Next i
    gSecCount = gSecCount - 1
    If gSecCount <= 0 Then
        Erase gSecStart
        Erase gSecEnd
        Erase gSecLevel
        Erase gSecTitle
        gLastIdx = 0
    Else
        ReDim Preserve gSecStart(1 To gSecCount)
        ReDim Preserve gSecEnd(1 To gSecCount)
        ReDim Preserve gSecLevel(1 To gSecCount)
        ReDim Preserve gSecTitle(1 To gSecCount)
        If gLastIdx > gSecCount Then gLastIdx = gSecCount
    End If
End Sub

Private Sub TrimOldestSections(ByVal removeCount As Long)
    Dim n As Long
    If removeCount <= 0 Then Exit Sub
    If removeCount >= gSecCount Then
        ResetSectionCache gDocKey, gFingerprint
        Exit Sub
    End If
    
    ' 맨 앞에서 removeCount개 제거
    Dim i As Long
    For i = 1 To gSecCount - removeCount
        gSecStart(i) = gSecStart(i + removeCount)
        gSecEnd(i) = gSecEnd(i + removeCount)
        gSecLevel(i) = gSecLevel(i + removeCount)
        gSecTitle(i) = gSecTitle(i + removeCount)
    Next i
    gSecCount = gSecCount - removeCount
    ReDim Preserve gSecStart(1 To gSecCount)
    ReDim Preserve gSecEnd(1 To gSecCount)
    ReDim Preserve gSecLevel(1 To gSecCount)
    ReDim Preserve gSecTitle(1 To gSecCount)
    gLastIdx = 0
End Sub

' =========================
' CustomXMLParts Persist/Load
' =========================

' 현재 메모리 캐시를 문서의 CustomXMLParts에 저장
Private Sub SaveCacheToDocument(ByVal doc As Document)
    On Error GoTo SafeExit
    
    ' 저장할 내용이 없으면 기존 파트도 정리
    If gSecCount <= 0 Then
        DeleteExistingOutlineParts doc
        Exit Sub
    End If
    
    Dim xml As String
    xml = BuildCacheXml(gDocKey, gFingerprint)
    
    ' 기존 파트 삭제 후 하나로 유지
    DeleteExistingOutlineParts doc
    doc.CustomXMLParts.Add xml
    Exit Sub
    
SafeExit:
    ' 저장 실패는 무시(문서 보호/권한/환경 차이 등)
End Sub

' 문서 내 저장된 캐시를 로드(유효하면 메모리 캐시에 적재)
' - fingerprint가 일치할 때만 로드
Private Function TryLoadCacheFromDocument( _
    ByVal doc As Document, _
    ByVal key As String, _
    ByVal fp As String _
) As Boolean
    On Error GoTo SafeExit
    
    Dim part As CustomXMLPart
    Set part = FindOutlinePart(doc)
    If part Is Nothing Then GoTo SafeExit
    
    Dim loadedFp As String
    Dim countLoaded As Long
    If Not ParseCacheXml(part.XML, loadedFp, countLoaded) Then GoTo SafeExit
    
    If loadedFp = "" Or loadedFp <> fp Then GoTo SafeExit
    
    ' ParseCacheXml가 전역 배열에 채워둠
    gDocKey = key
    gFingerprint = fp
    gLastIdx = 0
    
    TryLoadCacheFromDocument = True
    Exit Function
    
SafeExit:
    TryLoadCacheFromDocument = False
End Function

' 기존 Outline 캐시 XML 파트들 삭제
' (하나만 유지하기 위함)
Private Sub DeleteExistingOutlineParts( _
    ByVal doc As Document _
)
    On Error GoTo SafeExit
    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(OUTLINE_XML_NS)
    If Not parts Is Nothing Then
        Do While parts.Count > 0
            parts(1).Delete
        Loop
        Exit Sub
    End If
SafeExit:
    ' SelectByNamespace가 실패하는 환경이면 전체 순회로 삭제
    On Error Resume Next
    Dim p As CustomXMLPart
    For Each p In doc.CustomXMLParts
        If InStr(1, p.XML, OUTLINE_XML_NS, vbTextCompare) > 0 And _
            InStr(1, p.XML, OUTLINE_XML_ROOT_LOCAL, vbTextCompare) > 0 _
        Then
            p.Delete
        End If
    Next p
End Sub

' 우리 캐시 파트 1개 찾기
Private Function FindOutlinePart(ByVal doc As Document) As CustomXMLPart
    On Error GoTo Fallback
    Dim parts As CustomXMLParts
    Set parts = doc.CustomXMLParts.SelectByNamespace(OUTLINE_XML_NS)
    If Not parts Is Nothing Then
        If parts.Count > 0 Then
            Set FindOutlinePart = parts(1)
            Exit Function
        End If
    End If
    
Fallback:
    On Error Resume Next
    Dim p As CustomXMLPart
    For Each p In doc.CustomXMLParts
        If InStr(1, p.XML, OUTLINE_XML_NS, vbTextCompare) > 0 And _
            InStr(1, p.XML, OUTLINE_XML_ROOT_LOCAL, vbTextCompare) > 0 _
        Then
            Set FindOutlinePart = p
            Exit Function
        End If
    Next p
    Set FindOutlinePart = Nothing
End Function

' XML 문자열 생성
Private Function BuildCacheXml( _
    ByVal docKey As String, _
    ByVal fp As String _
) As String
    Dim sb As String
    Dim i As Long
    
    sb = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
         "<sl:" & OUTLINE_XML_ROOT_LOCAL & _
         " xmlns:sl=""" & OUTLINE_XML_NS & _
         """ version=""" & OUTLINE_XML_VERSION & """>" & _
         "<meta docKey=""" & EscapeXmlAttr(docKey) & _
         """ fingerprint=""" & EscapeXmlAttr(fp) & """ />" & _
         "<sections count=""" & CStr(gSecCount) & """>"
    
    For i = 1 To gSecCount
        sb = sb & "<sec s=""" & CStr(gSecStart(i)) & _
            """ e=""" & CStr(gSecEnd(i)) & _
            """ l=""" & CStr(gSecLevel(i)) & """>" & _
            "<t>" & EscapeXmlText(gSecTitle(i)) & "</t>" & _
            "</sec>"
    Next i
    
    sb = sb & "</sections></sl:" & OUTLINE_XML_ROOT_LOCAL & ">"
    BuildCacheXml = sb
End Function

' XML 파싱: 전역 캐시 배열로 로드하고 fingerprint/개수를 반환
Private Function ParseCacheXml( _
    ByVal xml As String, _
    ByRef outFingerprint As String, _
    ByRef outCount As Long _
) As Boolean
    On Error GoTo SafeExit
    
    outFingerprint = ""
    outCount = 0
    
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    
    If Not dom.LoadXML(xml) Then GoTo SafeExit
    
    Dim meta As Object
    Set meta = dom.selectSingleNode("//*[local-name()='meta']")
    If Not meta Is Nothing Then
        On Error Resume Next
        outFingerprint = CStr(meta.Attributes.getNamedItem("fingerprint").Text)
        On Error GoTo SafeExit
    End If
    
    Dim secNodes As Object
    Set secNodes = dom.selectNodes("//*[local-name()='sec']")
    If secNodes Is Nothing Then GoTo SafeExit
    
    Dim n As Long
    n = secNodes.Length
    If n <= 0 Then
        ResetSectionCache gDocKey, gFingerprint
        ParseCacheXml = True
        Exit Function
    End If
    
    If n > MAX_SECTION_CACHE Then n = MAX_SECTION_CACHE
    
    gSecCount = n
    ReDim gSecStart(1 To gSecCount)
    ReDim gSecEnd(1 To gSecCount)
    ReDim gSecLevel(1 To gSecCount)
    ReDim gSecTitle(1 To gSecCount)
    
    Dim i As Long
    For i = 1 To gSecCount
        Dim node As Object
        Set node = secNodes.Item(i - 1)
        
        gSecStart(i) = CLng(node.Attributes.getNamedItem("s").Text)
        gSecEnd(i) = CLng(node.Attributes.getNamedItem("e").Text)
        gSecLevel(i) = CLng(node.Attributes.getNamedItem("l").Text)
        
        Dim tNode As Object
        Set tNode = node.selectSingleNode("*[local-name()='t']")
        If Not tNode Is Nothing Then
            gSecTitle(i) = CStr(tNode.Text)
        Else
            gSecTitle(i) = ""
        End If
    Next i
    
    outCount = gSecCount
    ParseCacheXml = True
    Exit Function
    
SafeExit:
    ParseCacheXml = False
End Function

' XML attribute escape
Private Function EscapeXmlAttr(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscapeXmlAttr = s
End Function

' XML text escape
Private Function EscapeXmlText(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscapeXmlText = s
End Function

Private Function FormatHeadingTitle( _
    ByVal title As String, _
    ByVal maxLen As Long _
) As String
    Dim t As String
    t = title
    If maxLen > 0 And Len(t) > maxLen Then t = Left$(t, maxLen) & "…"
    If t = "" Then
        FormatHeadingTitle = ""
    Else
        FormatHeadingTitle = t
    End If
End Function

' 문서 구분용 key
Private Function MakeDocKey(ByVal doc As Document) As String
    On Error GoTo SafeExit
    MakeDocKey = doc.FullName
    Exit Function
SafeExit:
    MakeDocKey = doc.Name
End Function

' 문서 변경 여부 판단용 fingerprint (가볍게)
Private Function MakeFingerprint(ByVal doc As Document) As String
    On Error GoTo SafeExit
    MakeFingerprint = CStr(doc.Content.End)
    Exit Function
SafeExit:
    MakeFingerprint = ""
End Function

' 제목/문단 텍스트 정규화 (한 줄로)
Private Function NormalizeInlineText(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeInlineText = Trim$(s)
End Function

' 문단에 표시되는 "목록 번호(예: 1.2.3)" 문자열 반환
' - 다단계 번호/헤딩 번호 포함
' - 번호가 없거나 접근 실패 시 "" 반환
Private Function GetParagraphListNumber(ByVal p As Paragraph) As String
    On Error Resume Next
    Dim s As String
    s = p.Range.ListFormat.ListString
    If Err.Number <> 0 Then
        Err.Clear
        s = ""
    End If
    On Error GoTo 0
    GetParagraphListNumber = Trim$(s)
End Function

