export { };

addEvent({ elemID: "pretty-codec-text", event: "click", cb: () => toPrettyCodecText() });

// 밝은 보라색:#C04DFF
// 흐린 하늘색: #4F81BD
// 밝은 초록색: #9BBB59


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

    // const NoHead_az09 = "(?<![a-z0-9])";
    // const NoHead_AZ09 = "(?<![A-Z0-9])";
    // const NoHead_azAZ09 = "(?<![a-zA-Z0-9])";
    // const NoTail_azAZ09 = "(?!\w)";
    const searchMaps: SearchMap[] = [
      // 정규식 색칠
      { regex: /(?<!\w)[a-z][a-z0-9]*(_[a-z0-9]+)+/g, color: "#4F81BD" }, // snake_case
      { regex: /(?<!\w)[a-z][a-z0-9]*([A-Z][a-z0-9]*)+/g, color: "#00B050" }, // camelCase
      { regex: /(?<!\w)[A-Z][a-z0-9]+([A-Z][a-z0-9]*)+/g, color: "#0099FF" }, // PascalCase
      { regex: /(?<!\w)[A-Z][A-Z0-9]*(_[A-Z0-9]+)+/g, color: "#0066FF" }, // SCREAMING_SNAKE_CASE
      { regex: /(?<!\w)[a-zA-Z][_a-zA-Z0-9]*(?=\(.*?\))/g, color: "#F79646" }, // 함수명
      // { isRegex: true, search: "(?<![a-zA-Z0-9])[a-z][0-9](?![a-zA-Z0-9])", color: "#FF0000" }, // 계수들 (ex. c1, c2, c3, ...)
      // 조건식
      // { search: "If", replacement: "If", highlight: "lightgray" },
      // { search: "When", replacement: "If", highlight: "lightgray" },
      // { search: "Otherwise", replacement: "Else", highlight: "lightgray" },
      // { search: "until", replacement: "until", highlight: "lightgray" },
      // 비교문/할당문
      // { search: "is equal to", replacement: "= ="},
      // { search: "is not equal to", replacement: "!="},
      // { search: "is greater than or equal to", replacement: "≥"}, // 순서주의 1
      // { search: "is greater than", replacement: ">"},             // 순서주의 2
      // { search: "is less than or equal to", replacement: "≤"}, // 순서주의 1
      // { search: "is less than", replacement: "<"},             // 순서주의 2
      // { search: "is set equal to", replacement: ":="},   // 순서주의 1
      // { search: "are set equal to", replacement: "::="}, // 순서주의 2
      // { search: "set equal to", replacement: "="},       // 순서주의 3
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
      { search: "구문", replacement: "신택스" },
      { search: "구성 요소", replacement: "성분" },
      { search: "휘도", replacement: "루마" },
      { search: "명도", replacement: "루마" },
      { search: "루미너스", replacement: "루마" },
      { search: "색채", replacement: "크로마" },
      { search: "색차", replacement: "크로마" },
      { search: "채도", replacement: "크로마" },
      { search: "목록", replacement: "리스트" },
      { search: "영상", replacement: "픽처" },
      { search: "계산 복잡도", replacement: "계산복잡도" },
      { search: "참조 샘플", replacement: "참조샘플" },
      { search: "기준 샘플", replacement: "참조샘플" },
      { search: "그라디언트", replacement: "그래디언트" },
      { search: "화면 콘텐츠", replacement: "스크린 콘텐츠" },

      { search: "있습니다", replacement: "있다" },
      { search: "됩니다", replacement: "된다" },
      { search: "되었습니다", replacement: "되었다" },
      { search: "입니다", replacement: "이다" },
      { search: "합니다", replacement: "한다" },
      // 예외: 동일합니다 -> 동일'하다'
      // 예외: 필요합니다 -> 필요'하다'
      { search: "않습니다", replacement: "않는다" },
      { search: "줍니다", replacement: "준다" },
      { search: "나타냅니다", replacement: "나타낸다" },
      { search: "같습니다", replacement: "같다" },
      { search: "없습니다", replacement: "없다" },

      { search: "신호된다", replacement: "시그널링된다" },
      { search: "신호화된다", replacement: "시그널링된다" },
      { search: "신호화되어", replacement: "시그널링되어" },
      { search: "신호로 전달된다", replacement: "시그널링된다" },
      { search: "신호로 전달하여", replacement: "시그널링하여" },
      { search: "신호로 전송된다", replacement: "시그널링된다" },
      { search: "신호로 전송하여", replacement: "시그널링하여" },
      { search: "신호화가", replacement: "시그널링이" },
      { search: "신호 전송 시", replacement: "시그널링 시" },
      { search: "플래그가 전송", replacement: "플래그가 시그널링" },

      { search: "“", replacement: "\"" },
      { search: "”", replacement: "\"" },

      { search: "교차 구성", replacement: "교차 성분" }, // cross-component
      { search: "교차 구성 요소", replacement: "교차 성분" }, // cross-component
      { search: "부분 샘플링", replacement: "서브샘플링" }, // sub-sampling
      { search: "다중 모델", replacement: "멀티 모델" }, // multi-model
      { search: "사용 가능한", replacement: "가용한" }, // available
      { search: "사용 불가능한", replacement: "비가용한" }, // unavailable
      { search: "도출", replacement: "유도" }, // derivation
      { search: "자동상관", replacement: "자기상관" }, // autocorrelation

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
      const results = searchRange.search(searchText.search!);

      results.load("items");
      await context.sync();
      if (results.items.length === 0) {
        continue;
      }

      searchText.matchedRanges = results.items;
    }
  }

  // 검색 결과 꾸미기 (일반 검색, 정규식 검색 모두 포함)
  for (const { matchedRanges, color, highlight, replacement } of searchMap) {
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
      if (replacement) {
        item.insertText(replacement, "Replace");
      }
    }
  }
  await context.sync();
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

    // 현재 bold 상태를 확인
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
