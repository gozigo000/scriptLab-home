/// <reference path="../../../index.d.ts" />

document.getElementById("register-event-handler")!.addEventListener("click", () => tryCatch(registerEventHandler));
document.getElementById("get-paragraph-by-id")!.addEventListener("click", () => tryCatch(getParagraphById));
document.getElementById("deregister-event-handler")!.addEventListener("click", () => tryCatch(deregisterEventHandler));
document.getElementById("setup")!.addEventListener("click", () => tryCatch(setup));

let eventContext;

async function registerEventHandler() {
  // Registers the onParagraphChanged event handler on the document.
  await Word.run(async (context) => {
    eventContext = context.document.onParagraphChanged.add(paragraphChanged);
    await context.sync();

    console.log("Added event handler for when content is changed in paragraphs.");
  });
}

async function paragraphChanged(event: Word.ParagraphChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs where content was changed:`, event.uniqueLocalIds);
  });
}

async function getParagraphById() {
  await Word.run(async (context) => {
    const paragraphId = (document.getElementById("paragraph-id") as HTMLInputElement).value;
    const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
    paragraph.load();
    await paragraph.context.sync();

    console.log(paragraph);
  });
}

async function deregisterEventHandler() {
  await Word.run(eventContext.context, async (context) => {
    eventContext.remove();
    await context.sync();
  });

  eventContext = null;
  console.log("Removed event handler that was tracking content changes in paragraphs.");
}

async function setup() {
  await Word.run(async (context) => {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph("Paragraph 1", "End");
    body.insertParagraph("Paragraph 2", "End");
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
