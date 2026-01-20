// 파일을 모듈로 만들기 위한 빈 export
export { };

document.getElementById("setup")!
  .addEventListener("click", () => tryCatch(setup));
document.getElementById("before")!
  .addEventListener("click", () => tryCatch(before));
document.getElementById("start")!
  .addEventListener("click", () => tryCatch(start));
document.getElementById("end")!
  .addEventListener("click", () => tryCatch(end));
document.getElementById("after")!
  .addEventListener("click", () => tryCatch(after));
document.getElementById("replace")!
  .addEventListener("click", () => tryCatch(replace));

/**
 * 첫 번째 문단 위에 새로운 문단을 삽입합니다.
 */
async function before() {
  await Word.run(async (context) => {
    const range: Word.Paragraph = context.document.body.paragraphs
      .getFirst()
      .insertParagraph("This is Before", "Before");
    range.font.highlightColor = "yellow";

    await context.sync();
  });
}

/**
 * 첫 번째 문단 아래에 새로운 문단을 삽입합니다.
 */
async function after() {
  await Word.run(async (context) => {
    // Insert a paragraph after an existing one.
    const range: Word.Paragraph = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertParagraph("This is After", "After");
    range.font.highlightColor = "red";
    range.font.color = "white";

    await context.sync();
  });
}

/**
 * 첫 번째 문단의 시작 위치에 텍스트를 삽입합니다.
 */
async function start() {
  await Word.run(async (context) => {
    // This button assumes before() ran.
    // Get the next paragraph and insert text at the beginning. Note that there are invalid locations depending on the object. For instance, insertParagraph and "before" on a paragraph object is not a valid combination.
    const range: Word.Range = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertText("This is Start", "Start");
    range.font.highlightColor = "blue";
    range.font.color = "white";

    await context.sync();
  });
}

/**
 * 첫 번째 문단의 끝 위치에 텍스트를 삽입합니다.
 */
async function end() {
  await Word.run(async (context) => {
    // Insert text at the end of a paragraph.
    const range: Word.Range = context.document.body.paragraphs
      .getFirst()
      .getNext()
      .insertText(" This is End", "End");
    range.font.highlightColor = "green";
    range.font.color = "white";

    await context.sync();
  });
}

/**
 * 마지막 문단을 대체합니다.
 */
async function replace() {
  await Word.run(async (context) => {
    // Replace the last paragraph.
    const range: Word.Range = context.document.body.paragraphs
      .getLast()
      .insertText("Just replaced the last paragraph!", "Replace");
    range.font.highlightColor = "black";
    range.font.color = "white";

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
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
