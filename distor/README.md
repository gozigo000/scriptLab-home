# distor 도구 안내

이 폴더는 `myVba` 소스(`*.vba/.cls/.frm`)를 Word VBA 프로젝트 구조에 맞게 `dist/`로 변환하고,
그 결과물을 **Normal.dotm**에 일괄 반영하기 위한 스크립트/모듈을 제공합니다.

## 1) 변환: `myVba` → `dist` (PowerShell)

`distor/ConvertMyVbaToDist.ps1` 를 실행하면 `myVba` 아래의 소스를 찾아 Word VBA 프로젝트 구조에 맞춰 `dist/` 아래로 생성/정리합니다.

- **`dist/MsWord객체/`**
  - `ThisDocument.cls` (문서 모듈용 코드 파일: `Attribute VB_*` 헤더 없이 “코드만” 들어있음)
- **`dist/모듈/`**
  - 표준 모듈 `*.bas` (필요 시 `Attribute VB_Name` 헤더가 포함됨)
- **`dist/클래스모듈/`**
  - 클래스 모듈 `*.cls` (VBE Import 가능한 export 헤더가 포함되도록 래핑)
  - 단, `ThisDocument.cls`는 예외로 **클래스모듈이 아니라 `MsWord객체`로 분류**
- **`dist/폼/`**
  - 폼 `*.frm` (있으면 `.frx`도 함께 복사)

예시:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\distor\ConvertMyVbaToDist.ps1
```

옵션:
- `-InPlace`: 원본 `.vba` 파일의 확장자를 **그대로 `.bas`로 변경(리네임)** 합니다. (권장하지 않음)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\distor\ConvertMyVbaToDist.ps1 -InPlace
```

## 2) 반영: `dist` → `Normal.dotm` (VBA 매크로)

`distor/ImportMyVbaToNormal.bas` 를 VBA 편집기(Alt+F11)에서 표준 모듈로 가져온 다음,
`ImportMyVbaToNormal` 매크로를 실행하세요.

사전 조건(Word 보안 설정):
- Word → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정
  - **“VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스”** 체크

사용 방법:
- `distor/ImportMyVbaToNormal.bas`의 `rootDir`를 본인 PC 경로로 수정
```vba
Private Const rootDir As String = "C:\Users\gozig\Desktop\scriptLab-home"
```

동작:
- **먼저 변환 스크립트를 실행**: `ConvertMyVbaToDist.ps1`를 PowerShell로 실행(완료까지 대기)
- **그 다음 Import 수행**:
  - 동일한 `VB_Name`(또는 파일명 기반 모듈명)가 이미 Normal.dotm에 있으면 **삭제(Remove) 후 Import**
  - `ThisDocument`는 삭제할 수 없으므로, `dist/MsWord객체/ThisDocument.cls` 내용을 **ThisDocument 코드로 교체**
