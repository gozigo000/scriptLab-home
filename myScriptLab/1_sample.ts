export { };

addEvent({ elemID: "pretty-codec-text", event: "click", cb: () => toPrettyCodecText() });

// 밝은 보라색:#C04DFF
// 흐린 하늘색: #4F81BD
// 밝은 초록색: #9BBB59


// (MARK) 표준문서 꾸미기 섹션
// ----------------------
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

    // snake_case
    await _paintWildcard("<[a-z0-9]{1,}_[a-z0-9_]{1,}>", "#0099FF", selection);
    // camelCase
    await _paintWildcard("<[a-z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", "#0099FF", selection); // 예: log2CbSizeC
    await _paintWildcard("<[a-z][A-Z][a-zA-Z0-9]{1,}>", "#0099FF", selection); // 예: xCb, nCbSX, xTbCmp, cIdx
    await _paintWildcard("<[a-z]{1,}[A-Z0-9]>", "#0099FF", selection); // 예: xY, x0, yY, qP, bS
    // PascalCase
    await _paintWildcard("<[A-Z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", "#0099FF", selection);
    // SCREAMING_SNAKE_CASE
    await _paintWildcard("<[A-Z0-9]{1,}_[A-Za-z0-9_]{1,}>", "#0066FF", selection);
    // 비교문/할당문
    await _replaceSearch("is equal to", "= =", undefined, selection);
    await _replaceSearch("is not equal to", "!=", undefined, selection);
    await _replaceSearch("is greater than or equal to", "≥", undefined, selection); // 순서주의 1
    await _replaceSearch("is greater than", ">", undefined, selection);             // 순서주의 2
    await _replaceSearch("is less than or equal to", "≤", undefined, selection); // 순서주의 1
    await _replaceSearch("is less than", "<", undefined, selection);             // 순서주의 2
    await _replaceSearch("is set equal to", ":=", undefined, selection);  // 순서주의 1
    await _replaceSearch("are set equal to", ":=", undefined, selection); // 순서주의 2
    await _replaceSearch("set equal to", ":=", undefined, selection);     // 순서주의 3
    // acronyms
    await _replaceSearch("coding unit", "CU", undefined, selection);
    await _replaceSearch("coding block", "CB", undefined, selection);
    await _replaceSearch("block vector", "BV", undefined, selection);
    await _replaceSearch("motion vector", "MV", undefined, selection);
    // MY 약어
    await _replaceSearch("location", "loc", undefined, selection);
    await _replaceSearch("picture", "pic", undefined, selection);
    await _replaceSearch("variable", "var", undefined, selection);
    await _replaceSearch("current", "curr", undefined, selection);
    // 조건식
    await _replaceWildcard("<If>", "If", "lightgray", selection);
    await _replaceWildcard("<When>", "`If", "lightgray", selection);
    await _replaceSearch("Otherwise", "else", "lightgray", selection);
    await _replaceSearch("until", "until", "lightgray", selection);
    // await replaceSearch("and", "&&", "lightgray");
    // 곱하기
    await _replaceSearch(")x(", ")×(", undefined, selection);
    // 연한 회색 배경
    await _replaceSearch("is invoked with", "is invoked with", "lightgray", selection);
    await _replaceSearch("as outputs", "as outputs", "lightgray", selection);
    await _replaceSearch("as output", "as output", "lightgray", selection);
    await _replaceSearch("as inputs", "as inputs", "lightgray", selection);
    await _replaceSearch("as input", "as input", "lightgray", selection);
  });
}

/**
 * 와일드카드 색칠
 * @param wildcard - 와일드카드 표현식
 * @param color - 색상
 * @param searchRange - 검색할 범위 (선택 사항, 없으면 전체 문서)
 */
async function _paintWildcard(wildcard: string, color: string, searchRange?: Word.Range) {
  await Word.run(async (context) => {
    const searchTarget = searchRange || context.document.body;
    const rangeColl: Word.RangeCollection =
      searchTarget.search(wildcard, {
        matchWildcards: true
      });
    rangeColl.load("length");
    await context.sync();

    rangeColl.items.forEach(item => {
      item.font.color = color;
    });
    await context.sync();
  });
}

/**
 * 일반검색 바꾸기
 * @param searchText - 와일드카드 표현식
 * @param replacement - 바꿀 텍스트
 * @param highlight - 하이라이트 색상 (선택 사항)
 * @param searchRange - 검색할 범위 (선택 사항, 없으면 전체 문서)
 */
async function _replaceSearch(searchText: string, replacement: string, highlight?: string, searchRange?: Word.Range) {
  await Word.run(async (context) => {
    const searchTarget = searchRange || context.document.body;
    const rangeColl: Word.RangeCollection =
      searchTarget.search(searchText);
    rangeColl.load("length");
    await context.sync();

    // 검색된 각 항목을 replacement 텍스트로 교체 (역순으로 처리하여 위치 변경 문제 방지)
    const items = rangeColl.items;
    for (let i = items.length - 1; i >= 0; i--) {
      const item = items[i];
      item.insertText(replacement, "Replace");
      if (highlight) {
        item.font.highlightColor = highlight;
      }
    }
    await context.sync();
  });
}

/**
 * 와일드카드 바꾸기
 * @param wildcard - 와일드카드 표현식
 * @param replacement - 바꿀 텍스트
 * @param highlight - 하이라이트 색상 (선택 사항)
 * @param searchRange - 검색할 범위 (선택 사항, 없으면 전체 문서)
 */
async function _replaceWildcard(wildcard: string, replacement: string, highlight?: string, searchRange?: Word.Range) {
  await Word.run(async (context) => {
    const searchTarget = searchRange || context.document.body;
    const rangeColl: Word.RangeCollection =
      searchTarget.search(wildcard, {
        matchWildcards: true
      });
    rangeColl.load("length");
    await context.sync();

    // 검색된 각 항목을 replacement 텍스트로 교체 (역순으로 처리하여 위치 변경 문제 방지)
    const items = rangeColl.items;
    for (let i = items.length - 1; i >= 0; i--) {
      const item = items[i];
      item.insertText(replacement, "Replace");
      if (highlight) {
        item.font.highlightColor = highlight;
      }
    }
    await context.sync();
  });
}

// (MARK) 텍스트 치환 섹션
// ----------------------
// 텍스트 치환 관련 이벤트 등록
addEvent({ elemID: "replaceIsEqual", event: "click", cb: () => replaceIsEqual() });

/**
 * 선택 범위에서 "is equal to"를 "= ="로 바꾸는 함수
 */
async function replaceIsEqual() {
  await Word.run(async (context) => {
    // 선택 범위 가져오기
    const selection: Word.Range = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text.trim();
    if (!selectedText) {
      return;
    }

    // 선택 범위 내에서 "is equal to" 검색
    const results = selection.search("is equal to", {
      matchCase: false,
    });
    results.load("items");
    await context.sync();

    if (results.items.length === 0) {
      return;
    }

    // 역순으로 바꾸기 (텍스트 위치가 변경되므로 뒤에서부터 처리)
    for (let i = results.items.length - 1; i >= 0; i--) {
      const item = results.items[i];
      item.insertText("= =", Word.InsertLocation.replace);
    }
    await context.sync();
  });
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
