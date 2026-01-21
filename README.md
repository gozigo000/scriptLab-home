# VBA 파일을 나의 Word 파일에 적용하기

## 사전준비: Word 보안 설정
- Word → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정
  - **“VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스”** 체크

## 적용방법
1) `distor/ImportMyVbaToNormal.bas`를 VBA 편집기(Alt+F11)에서 표준 모듈로 가져오세요.

2) `distor/ImportMyVbaToNormal.bas`의 `rootDir`를 자기 PC 경로로 수정하세요.
```vba
Private Const rootDir As String = "자기 PC의 scriptLab-home 폴더 경로"
```

3) `ImportMyVbaToNormal` 매크로를 실행하세요.
