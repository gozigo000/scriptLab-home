export { };

addEvent({ elemID: "pretty-codec-text", event: "click", cb: () => toPrettyCodecText() });
addEvent({ elemID: "plain-codec-text", event: "click", cb: () => toPlainCodecText() });
addEvent({ elemID: "korean-word-codec", event: "click", cb: () => toKoreanWordCodec() });

// 폰트: 'Aptos Display', 'Calibri'
// 밝은 보라색: #C04DFF
// 흐린 하늘색: #4F81BD
// 밝은 초록색: #9BBB59
// 밝은 파란색: #0066FF
// 어둔 파란색: #1F497D
// 어둔 빨간색: #984806


// (MARK) 표준문서 꾸미기 섹션
// ----------------------
type SearchMap = {
  // 검색 속성
  search?: string,
  regex?: RegExp,
  matchedRanges?: Word.Range[],
  // 꾸미기 속성
  replacement?: string,
  color?: string,
  highlight?: string,
  underline?: string | Word.UnderlineType,
}

/**
 * 표준문서 꾸미기
 */
async function toPrettyCodecText() {
  await Word.run(async (context) => {
    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();
    if (!selectedText) {
      return;
    }

    const searchMaps: SearchMap[] = [
      // 곱셈기호
      { regex: /(?<=[\)\d])x(?=[\(\d])/g, replacement: "×" },

      // 변수명/상수명
      { regex: /(?<!\w)[xy][0-9]?(?!\w)/g, color: "#00B050" }, // 좌표들 (ex. x, y, x0, y1)
      { regex: /(?<!\w)[a-z][a-z0-9]*(_[a-z0-9]+)+/g, color: "#4F81BD" }, // snake_case
      { regex: /(?<!\w)[a-z][a-z0-9]*([A-Z][a-z0-9]*)+/g, color: "#00B050" }, // camelCase
      { regex: /(?<!\w)[A-Z][a-z0-9]+([A-Z][a-z0-9]*)+/g, color: "#0099CC" }, // PascalCase
      { regex: /(?<!\w)[A-Z][A-Z0-9]*(_[A-Z0-9]+)+/g, color: "#984806" }, // SCREAMING_SNAKE_CASE
      // 함수명
      { regex: /(?<!\w)[a-zA-Z][_a-zA-Z0-9]*(?=\(.*?\))/g, color: "#F79646" },
      // 조건식
      // { search: "If", replacement: "If", highlight: "lightgray" },
      // { search: "When", replacement: "If", highlight: "lightgray" },
      // { search: "Otherwise", replacement: "Else", highlight: "lightgray" },
      // { search: "until", replacement: "until", highlight: "lightgray" },
      // 비교문/할당문
      { regex: /(is |are )?(not )?equal to \w+/g, underline: "DottedHeavy" },
      { regex: /(is |are )?(greater|less|smaller|larger) than (or equal to )?\w+/g, underline: "DottedHeavy" },
      { regex: /not present/g, underline: "DottedHeavy" },
      { regex: /(is |are )?set equal to \w+/g, underline: "Double" },
      { regex: /(?<=inferred to )be equal to \w+/g, underline: "Double" },
      // Acronyms
      // { search: "coding unit", replacement: "CU"},
      // { search: "coding block", replacement: "CB"},
      // { search: "block vector", replacement: "BV"},
      // { search: "motion vector", replacement: "MV"},
      // My Abbreviations
      // { search: "location", replacement: "loc"},
      // { search: "picture", replacement: "pic"},
      // { search: "variable", replacement: "var"},
      // { search: "current", replacement: "curr"},
      // 하위 프로세스 호출 및 입출력
      // { search: "is invoked with", replacement: "is invoked with", highlight: "lightgray" },
      // { search: "as outputs", replacement: "as outputs", highlight: "lightgray" },
      // { search: "as output", replacement: "as output", highlight: "lightgray" },
      // { search: "as inputs", replacement: "as inputs", highlight: "lightgray" },
      // { search: "as input", replacement: "as input", highlight: "lightgray" },

      // 디플 번역 정제
      // { search: "구문", replacement: "신택스" },
      // { search: "구성 요소", replacement: "성분" },
      // { search: "휘도", replacement: "루마" },
      // { search: "명도", replacement: "루마" },
      // { search: "루미너스", replacement: "루마" },
      // { search: "색채", replacement: "크로마" },
      // { search: "색차", replacement: "크로마" },
      // { search: "채도", replacement: "크로마" },
      // { search: "목록", replacement: "리스트" },
      // { search: "영상", replacement: "픽처" },
      // { search: "계산 복잡도", replacement: "계산복잡도" },
      // { search: "참조 샘플", replacement: "참조샘플" },
      // { search: "기준 샘플", replacement: "참조샘플" },
      // { search: "그라디언트", replacement: "그래디언트" },
      // { search: "화면 콘텐츠", replacement: "스크린 콘텐츠" },

      // { search: "있습니다", replacement: "있다" },
      // { search: "됩니다", replacement: "된다" },
      // { search: "되었습니다", replacement: "되었다" },
      // { search: "입니다", replacement: "이다" },
      // { search: "합니다", replacement: "한다" },
      // // 예외: 동일합니다 -> 동일'하다'
      // // 예외: 필요합니다 -> 필요'하다'
      // { search: "않습니다", replacement: "않는다" },
      // { search: "줍니다", replacement: "준다" },
      // { search: "나타냅니다", replacement: "나타낸다" },
      // { search: "같습니다", replacement: "같다" },
      // { search: "없습니다", replacement: "없다" },

      // { search: "신호된다", replacement: "시그널링된다" },
      // { search: "신호화된다", replacement: "시그널링된다" },
      // { search: "신호화되어", replacement: "시그널링되어" },
      // { search: "신호로 전달된다", replacement: "시그널링된다" },
      // { search: "신호로 전달하여", replacement: "시그널링하여" },
      // { search: "신호로 전송된다", replacement: "시그널링된다" },
      // { search: "신호로 전송하여", replacement: "시그널링하여" },
      // { search: "신호화가", replacement: "시그널링이" },
      // { search: "신호 전송 시", replacement: "시그널링 시" },
      // { search: "플래그가 전송", replacement: "플래그가 시그널링" },

      // { search: "“", replacement: "\"" },
      // { search: "”", replacement: "\"" },

      // { search: "top", replacement: "top (↑)" },
      // { search: "above", replacement: "above (↑)" },
      // { search: "bottom", replacement: "bottom (↓)" },
      // { search: "below", replacement: "below (↓)" },
      // { search: "left", replacement: "left (←)" },
      // { search: "right", replacement: "right (→)" },
      // { search: "center", replacement: "center ()" },
      // { search: "middle", replacement: "middle ()" },
      // { search: "start", replacement: "start ()" },
      // { search: "end", replacement: "end ()" },

    ];
    await _reformatSearch(searchMaps, selection, context);
  });
}

/**
 * 검색 결과 꾸미기
 */
async function _reformatSearch(searchMap: SearchMap[], searchRange: Word.Range, context: Word.RequestContext) {
  searchRange.load("text");
  await context.sync();
  const rangeText = searchRange.text;

  // searchMap 속성 채우기
  for (const searchText of searchMap) {
    if (searchText.regex) {
      // 정규식 검색 처리
      // 정규식으로 매칭되는 모든 텍스트 찾기
      const matchedTexts: string[] = [];
      let match: RegExpExecArray | null;
      while ((match = searchText.regex.exec(rangeText)) !== null) {
        const matchedText = match[0];
        if (matchedText && matchedTexts.indexOf(matchedText) === -1) {
          matchedTexts.push(matchedText);
        }
      }

      // 각 매칭된 텍스트를 검색하여 Range 배열에 수집
      const matchedRanges: Word.Range[] = [];
      for (const matchedText of matchedTexts) {
        // 특수문자가 포함된 경우 이스케이프 처리
        const escapedText = matchedText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const results = searchRange.search(escapedText, {
          matchCase: true,
          matchWholeWord: true,
        });
        results.load("items");
        await context.sync();
        matchedRanges.push(...results.items);
      }

      if (matchedRanges.length === 0) {
        continue;
      }

      searchText.matchedRanges = matchedRanges;

    } else {
      // 일반 검색 처리
      const results = searchRange.search(searchText.search!, {
        matchCase: true,
        matchPrefix: true,
      });

      results.load("items");
      await context.sync();
      if (results.items.length === 0) {
        continue;
      }

      searchText.matchedRanges = results.items;
    }
  }

  // 검색 결과 꾸미기 (일반 검색, 정규식 검색 모두 포함)
  for (const { matchedRanges, replacement, color, highlight, underline } of searchMap) {
    if (!matchedRanges) {
      continue;
    }

    // 검색 결과 처리
    for (let i = matchedRanges.length - 1; i >= 0; i--) {
      const item = matchedRanges[i];
      if (color) {
        item.font.color = color;
      }
      if (highlight) {
        item.font.highlightColor = highlight;
      }
      if (underline) {
        item.font.underline = underline as Word.UnderlineType;
      }
      if (replacement) {
        item.insertText(replacement, "Replace");
      }
    }
  }
  await context.sync();
}

/**
 * 단어: 영어 -> 국어 변환
 */
async function toKoreanWordCodec() {
  await Word.run(async (context) => {
    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();
    if (!selectedText) {
      return;
    }

    const searchMaps: SearchMap[] = [
      // 단어: 영어 -> 국어
      { search: "sample", replacement: "샘플"},
      { search: "luma", replacement: "루마" },
      { search: "chroma", replacement: "크로마" },
      { search: "predictor", replacement: "예측기"},
      { search: "mode", replacement: "모드"},
      { search: "current", replacement: "현재"},
      { search: "previous", replacement: "이전"},
      { search: "next", replacement: "다음"},
      { search: "index", replacement: "인덱스"},    // TODO: 역변환시 단수형/복수형 고려해야 됨
      { search: "indices", replacement: "인덱스s"}, // 인덱스 -> index / 인덱스s -> indices
      { search: "vertical", replacement: "수직"},
      { search: "horizontal", replacement: "수평"},
      // '팔레트 모드' 관련 단어
      { search: "palette", replacement: "팔레트"},
    ];
    await _reformatSearch(searchMaps, selection, context);
  });
}

/**
 * 표준문서 꾸미기 제거 (선택 범위의 모든 꾸미기 초기화)
 */
async function toPlainCodecText() {
  await Word.run(async (context) => {
    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();

    // 선택 범위 내 모든 글자의 꾸미기 제거
    selection.font.color = "#000000"; // 글자색 제거
    selection.font.highlightColor = ""; // 하이라이트 제거
    selection.font.underline = "None"; // 밑줄 제거
    await context.sync();
  });
}


// (MARK) Bold 섹션
// ----------------------
// Bold 관련 이벤트 등록
addEvent({ elemID: "toggleBold", event: "click", cb: () => toggleBold() });

// 선택 범위의 텍스트 Bold 토글
async function toggleBold() {
  await Word.run(async (context) => {
    // 선택범위에 있는 텍스트 얻기
    const selection: Word.Range = context.document.getSelection();
    selection.load(["text", "font"]);
    await context.sync();

    const selectedText = selection.text.trim();
    if (!selectedText) {
      // console.log("선택된 텍스트가 없습니다.");
      return;
    }

    // 현재 선택 범위의 bold 상태 확인
    const isBold = selection.font.bold;

    // 선택된 텍스트를 검색
    const results = context.document.body.search(selectedText, {
      matchCase: true,
    });
    results.load("items");
    await context.sync();

    results.items.forEach((item) => {
      item.font.bold = !isBold;
    });
    await context.sync();
  });
}



// (MARK) 하이라이트 섹션
// ----------------------
// 하이라이트 색상 목록
const colors = [
  "Yellow",
  "Lime",
  "Turquoise",
  "Pink",
  "Blue",
  "Red",
  "DarkBlue",
  "Teal",
  "Green",
  "Purple",
  "DarkRed",
  "Olive",
  "Gray",
  "LightGray",
  "Black",
];

// 각 색상별 이벤트 등록
colors.forEach((color) => {
  const highlightId = `highlight${color}`;
  addEvent({ elemID: highlightId, event: "click", cb: () => toggleHighlight(color) });
});

/**
 * 하이라이트 버튼 클릭 시 하이라이트 기능을 토글하는 함수
 * @param color 하이라이트 색상
 */
async function toggleHighlight(color: string) {
  const buttonId = `highlight${color}`;
  const button = document.getElementById(buttonId);
  if (!button) return;
  const labelSpan = button.querySelector(".ms-Button-label");
  if (!labelSpan) return;

  const savedWord = labelSpan.textContent.trim();
  const isDefaultText = !savedWord || savedWord === "(강조)";

  await Word.run(async (context) => {
    // 선택범위에 있는 텍스트 얻기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();
    const selectedText = selection.text.trim();

    // 경우 1: 선택영역이 없고, 버튼에 저장된 단어가 있으면 → 리셋
    if (!selectedText && !isDefaultText) {
      const results = context.document.body.search(savedWord);
      results.load("items");
      await context.sync();

      // 해당 단어의 하이라이트 제거
      results.items.forEach((item) => {
        item.font.highlightColor = "";
      });
      await context.sync();

      // 버튼 텍스트를 원래대로 복원
      labelSpan.textContent = "";
      return;
    }

    // 경우 2: 선택영역이 있고, 버튼에 저장된 단어가 있으면 → 기존 강조 취소
    if (selectedText && !isDefaultText) {
      // 기존 단어의 하이라이트 제거
      const oldResults = context.document.body.search(savedWord);
      oldResults.load("items");
      await context.sync();

      oldResults.items.forEach((item) => {
        item.font.highlightColor = "";
      });
      await context.sync();
    }

    // 경우 3: 선택영역이 있으면 → 새로 강조 (또는 경우 2에서 기존 강조 취소 후)
    if (selectedText) {
      const results = context.document.body.search(selectedText);
      results.load("items");
      await context.sync();

      // 하이라이트 칠하기
      results.items.forEach((item) => {
        item.font.highlightColor = color;
      });
      await context.sync();

      // 버튼 텍스트를 강조된 단어로 변경
      labelSpan.textContent = selectedText;
    }
  });
}

// (MARK) 선택 영역 자동 검색 섹션
// ----------------------
// 페이지 로드 시 이벤트 핸들러 등록 (기본적으로 비활성화 상태)
// 사용자가 버튼을 클릭하면 활성화됨

let isSelectionHandlerRegistered = false;
let isSelectionAutoSearchEnabled = false;
let previousSelectedText: string | null = null;

// 버튼 클릭 이벤트 등록
addEvent({ elemID: "toggleSelectionAutoSearch", event: "click", cb: toggleSelectionAutoSearch });

/**
 * 선택 영역 자동 검색 기능 토글
 */
async function toggleSelectionAutoSearch() {
  const button = document.getElementById("toggleSelectionAutoSearch");
  const labelSpan = button!.querySelector(".ms-Button-label");
  
  if (!button || !labelSpan) {
    return;
  }

  isSelectionAutoSearchEnabled = !isSelectionAutoSearchEnabled;

  if (isSelectionAutoSearchEnabled) {
    // 기능 활성화
    button.classList.remove("inactive");
    button.classList.add("active");
    labelSpan.textContent = "비활성화";
    
    // 이벤트 핸들러가 등록되어 있지 않으면 등록
    if (!isSelectionHandlerRegistered) {
      registerSelectionChangedHandler();
    }
  } else {
    // 기능 비활성화
    button.classList.remove("active");
    button.classList.add("inactive");
    labelSpan.textContent = "활성화";
    
    // 이전 하이라이트 제거
    if (previousSelectedText) {
      Word.run(async (context) => {
        await removeHighlight(context, previousSelectedText!);
        previousSelectedText = null;
      });
    }
  }
}

/**
 * 선택 영역 변경 이벤트 핸들러 등록
 */
function registerSelectionChangedHandler() {
  if (isSelectionHandlerRegistered) {
    return; // 이미 등록되어 있으면 중복 등록 방지
  }

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    console.warn("Office.js가 로드되지 않았습니다. 이벤트 핸들러를 등록할 수 없습니다.");
    return;
  }

  // Office.js의 DocumentSelectionChanged 이벤트 등록
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged as any,
    (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        isSelectionHandlerRegistered = true;
      } else {
        console.error("이벤트 핸들러 등록 실패:", result.error);
      }
    }
  );
}

/**
 * 이전 하이라이트 제거 함수 (context를 받는 헬퍼 함수)
 */
async function removeHighlight(context: Word.RequestContext, text: string) {
  try {
    const results = context.document.body.search(text, {
      matchCase: true,
      matchWholeWord: false,
    });
    results.load("items");
    await context.sync();

    // 이전 하이라이트 제거
    results.items.forEach((item) => {
      item.font.highlightColor = "";
    });
    await context.sync();
  } catch (error) {
    console.error("하이라이트 제거 중 오류:", error);
  }
}

/**
 * 선택 영역이 변경될 때 호출되는 함수
 */
function onSelectionChanged(eventArgs: any) {
  if (!isSelectionAutoSearchEnabled) {
    return; // 기능이 비활성화되어 있으면 종료
  }

  Word.run(async (context) => {
    try {
      // 현재 선택 영역 가져오기
      const selection: Word.Range = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();
      if (!selectedText) {
        // 선택된 텍스트가 없으면 이전 하이라이트 제거
        if (previousSelectedText) {
          await removeHighlight(context, previousSelectedText);
          previousSelectedText = null;
        }
        return;
      }
      
      // 줄바꿈 기호가 있으면 바로 종료
      if (selectedText.includes("\n") || selectedText.includes("\r")) {
        return;
      }

      // 이전 텍스트와 동일하면 하이라이트 유지하고 종료
      if (selectedText === previousSelectedText) {
        return;
      }

      // 이전에 하이라이트된 텍스트가 있으면 하이라이트 제거
      if (previousSelectedText) {
        await removeHighlight(context, previousSelectedText);
      }

      // 문서 전체에서 선택된 텍스트와 동일한 텍스트 검색
      const results = context.document.body.search(selectedText, {
        matchCase: true,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      // 찾은 모든 텍스트에 초록색 하이라이트 적용
      results.items.forEach((item) => {
        item.font.highlightColor = "Lime";
      });
      await context.sync();

      // 현재 선택된 텍스트를 이전 텍스트로 저장
      previousSelectedText = selectedText;
    } catch (error) {
      console.error("선택 영역 검색 및 하이라이트 적용 중 오류:", error);
    }
  });
}


// (MARK) 유틸리티 함수
// ----------------------
async function addEvent({ elemID, event, cb }: { elemID: string, event: string, cb: () => Promise<void> }) {
  const elem = document.getElementById(elemID);
  if (!elem) return;
  elem.addEventListener(event, async () => {
    try {
      await cb();
    } catch (error) {
      console.error(error);
    }
  });
}
