// 파일을 모듈로 만들기 위한 빈 export
export { };

document.getElementById("setup")!
  .addEventListener("click", () => tryCatch(setup));
document.getElementById("basic-search")!
  .addEventListener("click", () => tryCatch(basicSearch));
document.getElementById("wildcard-search")!
  .addEventListener("click", () => tryCatch(wildcardSearch));

/**
 * 기본 텍스트 검색
 */
async function basicSearch() {
  // Does a basic text search and highlights matches in the document.
  await Word.run(async (context) => {
    const results: Word.RangeCollection = context.document.body.search("extend");
    results.load("length");

    await context.sync();

    // Let's traverse the search results and highlight matches.
    for (let i = 0; i < results.items.length; i++) {
      results.items[i].font.highlightColor = "yellow";
    }

    await context.sync();
  });
}

/**
 * 와일드카드 검색
 */
async function wildcardSearch() {
  // Does a wildcard search and highlights matches in the document.
  await Word.run(async (context) => {
    // Construct a wildcard expression and set matchWildcards to true in order to use wildcards.
    const results: Word.RangeCollection =
      context.document.body.search("$*.[0-9][0-9]", {
        matchWildcards: true
      });
    results.load("length");

    await context.sync();

    // Let's traverse the search results and highlight matches.
    for (let i = 0; i < results.items.length; i++) {
      results.items[i].font.highlightColor = "red";
      results.items[i].font.color = "white";
    }

    await context.sync();
  });
}

/**
 * 초기 세팅
 */
async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start",
    );
    body.paragraphs
      .getLast()
      .insertText(
        "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
        "Replace",
      );

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
