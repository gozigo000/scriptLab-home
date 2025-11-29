/// <reference path="../../index.d.ts" />

// 파일을 모듈로 만들기 위한 빈 export
export {};

/**
 * WordApiDesktop 1.3이 지원되는지 확인하는 함수
 * @returns WordApiDesktop 1.3이 지원되면 true, 그렇지 않으면 false
 */
function isWordApiDesktop13Supported(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
}

// 사용 예시
console.log("WordApiDesktop 1.3 지원 여부:", isWordApiDesktop13Supported());
