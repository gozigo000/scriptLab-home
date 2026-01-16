Option Explicit

' (MARK) 문서 개요(탐색창 제목) 캐시 빌더
' ----------------------
' Lazy 캐시:
' - 문서 전체를 한 번에 스캔하지 않습니다(초반 빌드 비용 제거).
' - 커서가 "다른 아웃라인 영역(제목 단위)"로 이동했을 때만,
'   현재 커서가 속한 영역의 제목/경계를 계산해서 캐시합니다.
'
' 캐시 범위 정의(섹션):
' - headingStart(현재 제목 문단 시작) <= pos < nextBoundaryStart
' - nextBoundaryStart는 "다음 제목 중, 현재 제목의 OutlineLevel보다 같거나 높은(<=) 레벨"의 시작 위치
'
' 즉, Heading2 아래 Heading3들은 같은 섹션으로 보고,
' 다음 Heading2 또는 Heading1이 나오면 섹션이 바뀐 것으로 간주합니다.
'
' 제목 판정 기준:
' - Paragraph.OutlineLevel <> wdOutlineLevelBodyText
'
' 사용 예:
' - title = GetNearestHeadingTitle(Selection.Range, 140)
' - title = GetNearestHeadingTitleLazy(Selection.Range, 140)
' - Call InvalidateOutlineCache() ' 필요 시 강제 갱신/초기화

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
Private gSecEnd() As Long          ' 섹션 끝(포함) = nextBoundaryStart-1
Private gSecLevel() As Long
Private gSecTitle() As String
Private gLastIdx As Long           ' 직전 히트 인덱스(로컬리티 최적화)

' ===== 문서 내 저장(CustomXMLParts) 설정 =====
Private Const OUTLINE_XML_NS As String = "urn:scriptlab:outlinecache:v1"
Private Const OUTLINE_XML_ROOT_LOCAL As String = "outlineCache"
Private Const OUTLINE_XML_VERSION As String = "1"

' ===== Public API =====

' 현재 커서가 속한 문단에서 위쪽으로 가장 가까운 "제목(탐색창에 표시되는 개요 수준)"을 찾아 반환
' - rng: 현재 커서가 속한 문단의 Range
' - maxLen: 반환 문자열의 최대 길이
' - 반환 예: "[수준 2] 서론"
' 하위호환용 이름: 내부적으로 lazy 캐시 사용
Public Function GetNearestHeadingTitle(ByVal rng As Range, Optional ByVal maxLen As Long = 140) As String
    GetNearestHeadingTitle = GetNearestHeadingTitleLazy(rng, maxLen)
End Function

' Lazy 방식: 커서가 다른 섹션으로 이동했을 때만 제목/경계를 갱신합니다.
Public Function GetNearestHeadingTitleLazy(ByVal rng As Range, Optional ByVal maxLen As Long = 140) As String
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
        If pos >= gSecStart(gLastIdx) And pos <= gSecEnd(gLastIdx) Then
            GetNearestHeadingTitleLazy = FormatHeadingTitle(gSecLevel(gLastIdx), gSecTitle(gLastIdx), maxLen)
            TouchRamCacheEntry gDocKey, gFingerprint
            Exit Function
        End If
    End If
    
    ' 2) 캐시된 섹션 구간에서 이진탐색
    Dim idx As Long
    idx = FindCachedSectionIndex(pos)
    If idx > 0 Then
        gLastIdx = idx
        GetNearestHeadingTitleLazy = FormatHeadingTitle(gSecLevel(idx), gSecTitle(idx), maxLen)
        TouchRamCacheEntry gDocKey, gFingerprint
        Exit Function
    End If
    
    ' 3) 캐시 미스: 현재 위치 섹션을 계산하고 캐시에 추가
    Dim headingStart As Long, nextBoundaryStart As Long, headingLevel As Long, headingTitle As String
    ComputeSectionForRange doc, rng, headingStart, nextBoundaryStart, headingLevel, headingTitle
    
    If headingStart > 0 Then
        idx = AddOrUpdateSectionCache(headingStart, nextBoundaryStart, headingLevel, headingTitle)
        gLastIdx = idx
        GetNearestHeadingTitleLazy = FormatHeadingTitle(headingLevel, headingTitle, maxLen)
        ' 캐시가 늘어났으면 문서에 저장(초기 전체 스캔이 없어서 빈도는 낮은 편)
        Call SaveCacheToDocument(doc)
        ' RAM 캐시도 최신 상태로 반영
        UpsertRamCacheFromCurrent
    Else
        GetNearestHeadingTitleLazy = ""
    End If
    Exit Function
    
SafeExit:
    GetNearestHeadingTitleLazy = ""
End Function

' 캐시를 강제로 무효화합니다. (다음 호출 시 재빌드됨)
Public Sub InvalidateOutlineCache(Optional ByVal doc As Document)
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

' 현재 문서(key/fp)의 캐시가 활성화되도록 보장:
' - RAM에 있으면 즉시 복구
' - 없으면 CustomXMLParts에서 로드(유효하면) 후 RAM에 저장
Private Sub EnsureActiveDocumentCache(ByVal doc As Document, ByVal key As String, ByVal fp As String)
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
Private Function LoadFromRamCache(ByVal key As String, ByVal fp As String) As Boolean
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

Private Function FindRamCacheIndex(ByVal key As String, ByVal fp As String) As Long
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
    ByRef nextBoundaryStart As Long, _
    ByRef headingLevel As Long, _
    ByRef headingTitle As String _
)
    On Error GoTo SafeExit
    
    headingStart = 0
    nextBoundaryStart = doc.Content.End + 1
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
            headingTitle = NormalizeInlineText(p.Range.Text)
            Exit Do
        End If
        On Error Resume Next
        Set p = p.Previous
        On Error GoTo SafeExit
    Loop
    
    If headingStart = 0 Then Exit Sub
    
    ' 2) 아래로 내려가며 섹션 경계(다음 동일/상위 레벨 제목) 찾기
    Dim q As Paragraph
    Set q = doc.Range(headingStart, headingStart).Paragraphs(1).Next
    
    Do While Not q Is Nothing
        lvl = q.OutlineLevel
        If lvl <> wdOutlineLevelBodyText Then
            If CLng(lvl) <= headingLevel Then
                nextBoundaryStart = q.Range.Start
                Exit Do
            End If
        End If
        On Error Resume Next
        Set q = q.Next
        On Error GoTo SafeExit
    Loop
    
    Exit Sub
    
SafeExit:
    headingStart = 0
    nextBoundaryStart = 0
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
    
    ' 섹션 시작 배열에서 pos 이하인 마지막 인덱스를 찾고, 그 구간에 포함되는지 확인
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

' 섹션(headingStart~nextBoundaryStart-1)을 캐시에 삽입/업데이트하고 인덱스를 반환
Private Function AddOrUpdateSectionCache( _
    ByVal headingStart As Long, _
    ByVal nextBoundaryStart As Long, _
    ByVal headingLevel As Long, _
    ByVal headingTitle As String _
) As Long
    Dim secStart As Long, secEnd As Long
    secStart = headingStart
    secEnd = nextBoundaryStart - 1
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
            ' 겹침: 더 넓은 end를 유지, title/level은 "최근 삽입(idx)" 것을 유지
            If gSecEnd(idx - 1) > gSecEnd(idx) Then gSecEnd(idx) = gSecEnd(idx - 1)
            DeleteSectionAt idx - 1
            idx = idx - 1
        End If
    End If
    
    ' 오른쪽과 겹치면 오른쪽 제거
    If idx < gSecCount Then
        If gSecEnd(idx) >= gSecStart(idx + 1) Then
            If gSecEnd(idx + 1) > gSecEnd(idx) Then gSecEnd(idx) = gSecEnd(idx + 1)
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
        Erase gSecStart: Erase gSecEnd: Erase gSecLevel: Erase gSecTitle
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
Private Function TryLoadCacheFromDocument(ByVal doc As Document, ByVal key As String, ByVal fp As String) As Boolean
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

' 기존 Outline 캐시 XML 파트들 삭제(하나만 유지하기 위함)
Private Sub DeleteExistingOutlineParts(ByVal doc As Document)
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
        If InStr(1, p.XML, OUTLINE_XML_NS, vbTextCompare) > 0 And InStr(1, p.XML, OUTLINE_XML_ROOT_LOCAL, vbTextCompare) > 0 Then
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
        If InStr(1, p.XML, OUTLINE_XML_NS, vbTextCompare) > 0 And InStr(1, p.XML, OUTLINE_XML_ROOT_LOCAL, vbTextCompare) > 0 Then
            Set FindOutlinePart = p
            Exit Function
        End If
    Next p
    Set FindOutlinePart = Nothing
End Function

' XML 문자열 생성
Private Function BuildCacheXml(ByVal docKey As String, ByVal fp As String) As String
    Dim sb As String
    Dim i As Long
    
    sb = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
         "<sl:" & OUTLINE_XML_ROOT_LOCAL & " xmlns:sl=""" & OUTLINE_XML_NS & """ version=""" & OUTLINE_XML_VERSION & """>" & _
         "<meta docKey=""" & EscapeXmlAttr(docKey) & """ fingerprint=""" & EscapeXmlAttr(fp) & """ />" & _
         "<sections count=""" & CStr(gSecCount) & """>"
    
    For i = 1 To gSecCount
        sb = sb & "<sec s=""" & CStr(gSecStart(i)) & """ e=""" & CStr(gSecEnd(i)) & """ l=""" & CStr(gSecLevel(i)) & """>" & _
                  "<t>" & EscapeXmlText(gSecTitle(i)) & "</t>" & _
                  "</sec>"
    Next i
    
    sb = sb & "</sections></sl:" & OUTLINE_XML_ROOT_LOCAL & ">"
    BuildCacheXml = sb
End Function

' XML 파싱: 전역 캐시 배열로 로드하고 fingerprint/개수를 반환
Private Function ParseCacheXml(ByVal xml As String, ByRef outFingerprint As String, ByRef outCount As Long) As Boolean
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

Private Function FormatHeadingTitle(ByVal level As Long, ByVal title As String, ByVal maxLen As Long) As String
    Dim t As String
    t = title
    If maxLen > 0 And Len(t) > maxLen Then t = Left$(t, maxLen) & "…"
    If t = "" Then
        FormatHeadingTitle = ""
    Else
        FormatHeadingTitle = "[수준 " & CStr(level) & "] " & t
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

