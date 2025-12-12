export { };

addEvent({ elemID: "pretty-codec-text", event: "click", cb: () => toPrettyCodecText() });

// 밝은 보라색:#C04DFF
// 흐린 하늘색: #4F81BD
// 밝은 초록색: #9BBB59


// (MARK) 표준문서 꾸미기 섹션
// ----------------------
type SearchMap = {
  search: string,
  isWildcard?: boolean,
  replacement?: string,
  color?: string,
  highlight?: string,
  rangeCollection?: Word.RangeCollection,
  length?: number,
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
      // 와일드카드 색칠
      { isWildcard: true, search: "<[a-z0-9]{1,}_[a-z0-9_]{1,}>", color: "#0099FF" }, // snake_case
      { isWildcard: true, search: "<[a-z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", color: "#0099FF" }, // camelCase
      { isWildcard: true, search: "<[a-z][A-Z][a-zA-Z0-9]{1,}>", color: "#0099FF" }, // camelCase
      { isWildcard: true, search: "<[a-z]{1,}[A-Z0-9]>", color: "#0099FF" }, // camelCase
      { isWildcard: true, search: "<[A-Z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", color: "#0099FF" }, // PascalCase
      { isWildcard: true, search: "<[A-Z0-9]{1,}_[A-Za-z0-9_]{1,}>", color: "#0066FF" }, // SCREAMING_SNAKE_CASE
      // 조건식
      { isWildcard: true, search: "<If>", replacement: "If", highlight: "lightgray" },
      { isWildcard: true, search: "<When>", replacement: "If", highlight: "lightgray" },
      { isWildcard: true, search: "<Otherwise>", replacement: "Else", highlight: "lightgray" },
      { isWildcard: true, search: "<until>", replacement: "until", highlight: "lightgray" },
      // 비교문/할당문
      { search: "is equal to", replacement: "= ="},
      { search: "is not equal to", replacement: "!="},
      { search: "is greater than or equal to", replacement: "≥"}, // 순서주의 1
      { search: "is greater than", replacement: ">"},             // 순서주의 2
      { search: "is less than or equal to", replacement: "≤"}, // 순서주의 1
      { search: "is less than", replacement: "<"},             // 순서주의 2
      { search: "is set equal to", replacement: ":="},  // 순서주의 1
      { search: "are set equal to", replacement: ":="}, // 순서주의 2
      { search: "set equal to", replacement: ":="},     // 순서주의 3
      // acronyms
      { search: "coding unit", replacement: "CU"},
      { search: "coding block", replacement: "CB"},
      { search: "block vector", replacement: "BV"},
      { search: "motion vector", replacement: "MV"},
      // My abbreviations
      { search: "location", replacement: "loc"},
      { search: "picture", replacement: "pic"},
      { search: "variable", replacement: "var"},
      { search: "current", replacement: "curr"},
      // 하위 프로세스 호출 및 입출력
      { search: "is invoked with", replacement: "is invoked with", highlight: "lightgray" },
      { search: "as outputs", replacement: "as outputs", highlight: "lightgray" },
      { search: "as output", replacement: "as output", highlight: "lightgray" },
      { search: "as inputs", replacement: "as inputs", highlight: "lightgray" },
      { search: "as input", replacement: "as input", highlight: "lightgray" },
    ];
    await _reformatSearch(searchMaps, selection, context);
  });
}

/**
 * 와일드카드 색칠
 * @param wildcard - 와일드카드 표현식
 * @param color - 색상
 * @param searchRange - 검색범위
 */
async function _reformatSearch(searchMap: SearchMap[], searchRange: Word.Range, context: Word.RequestContext) {
  // searchMap 속성 채우기
  for (const searchText of searchMap) {
    const rangeColl: Word.RangeCollection =
      searchRange.search(searchText.search, {
        matchWildcards: searchText.isWildcard
      });

    rangeColl.load("length");
    await context.sync();
    if (rangeColl.items.length === 0) {
      continue;
    }

    searchText.rangeCollection = rangeColl;
    searchText.length = rangeColl.items.length;
  }

  // 검색된 텍스트 꾸미기
  for (const { color, highlight, replacement, length, rangeCollection } of searchMap) {
    for (let i = length! - 1; i >= 0; i--) {
      const item = rangeCollection!.items[i];
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



// (MARK) 정규식 검색 섹션
// ----------------------
// 정규식 검색 관련 이벤트 등록
addEvent({ elemID: "regexSearch", event: "click", cb: () => regexSearchAndColor() });
addEvent({ elemID: "regexClear", event: "click", cb: () => toPlain() });

/**
 * 정규식으로 검색하여 빨간색으로 칠하는 함수 (선택 범위에서만)
 */
async function regexSearchAndColor() {
  const regexInput = document.getElementById("regexInput") as HTMLInputElement;
  if (!regexInput) return;

  const pattern = regexInput.value.trim();
  if (!pattern) {
    alert("정규식을 입력해주세요.");
    return;
  }

  let regex: RegExp;
  try {
    regex = new RegExp(pattern, "g");
  } catch (error) {
    alert("올바른 정규식이 아닙니다: " + (error as Error).message);
    return;
  }

  await Word.run(async (context) => {

    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();

    if (!selectedText) {
      return;
    }

    // 선택 범위 내에서 정규식으로 매칭되는 모든 텍스트 찾기
    const matches: string[] = [];
    let match: RegExpExecArray | null;
    while ((match = regex.exec(selectedText)) !== null) {
      const matchedText = match[0];
      if (matchedText && matches.indexOf(matchedText) === -1) {
        matches.push(matchedText);
      }
    }

    if (matches.length === 0) {
      alert("선택 범위에서 매칭되는 텍스트가 없습니다.");
      return;
    }

    // 선택 범위 내에서 각 매칭된 텍스트를 검색하여 빨간색으로 칠하기
    for (const matchedText of matches) {
      // 특수문자가 포함된 경우 이스케이프 처리
      const escapedText = matchedText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const results = selection.search(escapedText, {
        matchCase: false,
      });
      results.load("items");
      await context.sync();

      results.items.forEach((item) => {
        item.font.color = "#FF0000"; // 빨간색 (RGB: FF0000)
      });
      await context.sync();
    }
  });
}

/**
 * 선택 범위 내 모든 글자의 색상을 검정으로 하고 하이라이트를 제거하는 함수
 */
async function toPlain() {
  await Word.run(async (context) => {
    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();

    if (!selectedText) {
      return;
    }

    // 선택 범위 내 모든 글자의 색상을 검정으로 설정하고 하이라이트 제거
    selection.font.color = "#000000"; // 검정색
    selection.font.highlightColor = ""; // 하이라이트 제거
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

    // 현재 bold 상태를 확인
    const isBold = selection.font.bold;

    // 선택된 텍스트를 검색
    const results = context.document.body.search(selectedText, {
      // ignorePunct: false,
      // ignoreSpace: false,
      matchCase: true,
      // matchPrefix: false,
      // matchSuffix: false,
      // matchWholeWord: true,
      // matchWildcards: false,
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
