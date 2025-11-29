// 파일을 모듈로 만들기 위한 빈 export
export { };

document.getElementById("add-text")!
  .addEventListener("click", () => tryCatch(addFormattedText));
document.getElementById("add-paragraph")!
  .addEventListener("click", () => tryCatch(addFormattedParagraph));
document.getElementById("add-preset")!
  .addEventListener("click", () => tryCatch(addPreStyledFormattedText));

async function addFormattedText() {
  await Word.run(async (context) => {
    // Insert the sentence, then adjust the formatting.
    // Note that replace affects the calling object, in this case the entire document body.
    // A similar method can also be used at the range level.
    const sentence: Word.Range = context.document.body.insertText("This is some formatted text!", "Replace");
    sentence.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });

    await context.sync();
  });
}

async function addFormattedParagraph() {
  await Word.run(async (context) => {
    // Second sentence, let's insert it as a paragraph after the previously inserted one.
    const secondSentence: Word.Paragraph = context.document.body.insertParagraph(
      "This is the first text with a custom style.",
      "End",
    );
    secondSentence.font.set({
      bold: false,
      italic: true,
      name: "Berlin Sans FB",
      color: "blue",
      size: 30,
    });

    await context.sync();
  });
}

async function addPreStyledFormattedText() {
  await Word.run(async (context) => {
    const sentence: Word.Paragraph = context.document.body.insertParagraph("To be or not to be", "End");

    // Use styleBuiltIn to use an enumeration of existing styles. If your style is custom make sure to use: range.style = "name of your style";
    sentence.styleBuiltIn = Word.BuiltInStyleName.intenseReference;

    await context.sync();
  });
}
// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

