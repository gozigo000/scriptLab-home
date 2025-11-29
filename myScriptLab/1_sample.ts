export { };

document.getElementById("wildcard-search")!
  .addEventListener("click", () => tryCatch(prettifyCodecText));

// 밝은 보라색:#C04DFF
// 흐린 하늘색: #4F81BD
// 밝은 초록색: #9BBB59

/**
 * 와일드카드 검색
 */
async function prettifyCodecText() {
  // snake_case
  await paintWildcard("<[a-z0-9]{1,}_[a-z0-9_]{1,}>", "#0099FF");
  // camelCase
  await paintWildcard("<[a-z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", "#0099FF"); // 예: log2CbSizeC
  await paintWildcard("<[a-z][A-Z][a-zA-Z0-9]{1,}>", "#0099FF"); // 예: xCb, nCbSX, xTbCmp, cIdx
  await paintWildcard("<[a-z]{1,}[A-Z0-9]>", "#0099FF"); // 예: xY, x0, yY, qP, bS
  // PascalCase
  await paintWildcard("<[A-Z][a-z0-9]{1,}[A-Z][a-zA-Z0-9]{1,}>", "#0099FF");
  // SCREAMING_SNAKE_CASE
  await paintWildcard("<[A-Z0-9]{1,}_[A-Za-z0-9_]{1,}>", "#0066FF");
  // 비교문/할당문
  await replaceSearch("is equal to", "==");
  await replaceSearch("is not equal to", "!=");
  await replaceSearch("is greater than or equal to", "≥"); // 순서주의 1
  await replaceSearch("is greater than", ">");             // 순서주의 2
  await replaceSearch("is less than or equal to", "≤"); // 순서주의 1
  await replaceSearch("is less than", "<");             // 순서주의 2
  await replaceSearch("is set equal to", `'=`);  // 순서주의 1
  await replaceSearch("are set equal to", `"=`); // 순서주의 2
  await replaceSearch("set equal to", "=");      // 순서주의 3
  // acronyms
  await replaceSearch("coding unit", "CU");
  await replaceSearch("coding block", "CB");
  await replaceSearch("block vector", "BV");
  await replaceSearch("motion vector", "MV");
  // MY 약어
  await replaceSearch("location", "loc");
  await replaceSearch("picture", "pic");
  await replaceSearch("variable", "var");
  await replaceSearch("current", "curr");
  // 조건식
  await replaceWildcard("<If>", "If", "lightgray");
  await replaceWildcard("<When>", "`If", "lightgray");
  await replaceSearch("Otherwise", "else", "lightgray");
  await replaceSearch("until", "until", "lightgray");
  // await replaceSearch("and", "&&", "lightgray");
  // 곱하기
  await replaceSearch(")x(", ")×(");
  // 연한 회색 배경
  await replaceSearch("is invoked with", "is invoked with", "lightgray");
  await replaceSearch("as outputs", "as outputs", "lightgray");
  await replaceSearch("as output", "as output", "lightgray");
  await replaceSearch("as inputs", "as inputs", "lightgray");
  await replaceSearch("as input", "as input", "lightgray");
}

/**
 * 와일드카드 색칠
 * @param wildcard - 와일드카드 표현식
 * @param color - 색상
 */
async function paintWildcard(wildcard: string, color: string) {
  await Word.run(async (context) => {
    const rangeColl: Word.RangeCollection =
      context.document.body.search(wildcard, {
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
 */
async function replaceSearch(searchText: string, replacement: string, highlight?: string) {
  await Word.run(async (context) => {
    const rangeColl: Word.RangeCollection =
      context.document.body.search(searchText);
    rangeColl.load("length");
    await context.sync();

    // 검색된 각 항목을 replacement 텍스트로 교체
    rangeColl.items.forEach(item => {
      item.insertText(replacement, "Replace");
      highlight && (item.font.highlightColor = highlight);
    });
    await context.sync();
  });
}

/**
 * 와일드카드 바꾸기
 * @param wildcard - 와일드카드 표현식
 * @param replacement - 바꿀 텍스트
 */
async function replaceWildcard(wildcard: string, replacement: string, highlight?: string) {
  await Word.run(async (context) => {
    const rangeColl: Word.RangeCollection =
      context.document.body.search(wildcard, {
        matchWildcards: true
      });
    rangeColl.load("length");
    await context.sync();

    // 검색된 각 항목을 replacement 텍스트로 교체
    rangeColl.items.forEach(item => {
      item.insertText(replacement, "Replace");
      if (highlight) {
        item.font.highlightColor = highlight;
      }
    });
    await context.sync();
  });
}

/**
 * 오류 처리
 */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
