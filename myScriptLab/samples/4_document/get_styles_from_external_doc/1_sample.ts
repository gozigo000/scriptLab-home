// 파일을 모듈로 만들기 위한 빈 export
export { };

document.getElementById("file")!
  .addEventListener("change", getBase64);
document.getElementById("get-external-styles")!
  .addEventListener("click", () => tryCatch(getExternalStyles));

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

async function getExternalStyles() {
  // Gets style info from another document passed in as a Base64-encoded string.
  await Word.run(async (context) => {
    const retrievedStyles =
      context.application.retrieveStylesFromBase64(externalDocument);
    await context.sync();

    console.log("Styles from the other document:", retrievedStyles.value);
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
