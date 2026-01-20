// 파일을 모듈로 만들기 위한 빈 export
export { };

document.getElementById("file")!
  .addEventListener("change", getBase64);
document.getElementById("insert-document")!
  .addEventListener("click", () => tryCatch(insertDocument));
document.getElementById("insert-document-with-settings")!
  .addEventListener("click", () => tryCatch(insertDocumentWithSettings));

let externalDocument: string;

function getBase64() {
  // Retrieve the file and set up an HTML FileReader element.
  const myFile = document.getElementById("file") as HTMLInputElement;
  const reader = new FileReader();

  reader.onload = (event) => {
    // Remove the metadata before the Base64-encoded string.
    const startIndex = reader.result!.toString().indexOf("base64,");
    externalDocument = reader.result!.toString().substr(startIndex + 7);
  };

  // Read the file as a data URL so that we can parse the Base64-encoded string.
  reader.readAsDataURL(myFile.files![0]);
}

/**
 * 외부 문서의 텍스트를 현재 문서에 삽입합니다. \
 * 외부 문서는 Base64-encoded string으로 전달됩니다.
 */
async function insertDocument() {
  await Word.run(async (context) => {
    if (!checkRequirementSet()) return;

    // Use the Base64-encoded string representation of the selected '.docx' file.
    const externalDoc: Word.DocumentCreated =
      context.application.createDocument(externalDocument);
    await context.sync();

    const externalDocBody: Word.Body = externalDoc.body;
    externalDocBody.load("text");
    await context.sync();

    // 외부 문서의 텍스트를 현재 문서의 본문 시작 부분에 삽입합니다.
    const externalDocBodyText = externalDocBody.text;
    const currentDocBody: Word.Body = context.document.body;
    currentDocBody.insertText(externalDocBodyText, "Start");
    await context.sync();
  });
}

/**
 * 외부 문서를 삽입합니다. (설정 적용)
 */
async function insertDocumentWithSettings() {
  // Inserts content (applying selected settings) from another document passed in as a Base64-encoded string.
  await Word.run(async (context) => {
    // Use the Base64-encoded string representation of the selected .docx file.
    context.document.insertFileFromBase64(externalDocument, "Replace", {
      importTheme: true,
      importStyles: true,
      importParagraphSpacing: true,
      importPageColor: true,
      importChangeTrackingMode: true,
      importCustomProperties: true,
      importCustomXmlParts: true,
      importDifferentOddEvenPages: true,
    });
    await context.sync();
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

function checkRequirementSet(): boolean {
  if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
    console.warn(
      "The WordApiHiddenDocument 1.3 requirement set isn't supported on this client so can't proceed. Try this action on a platform that supports this requirement set.",
    );
    return false;
  }

  return true;
}