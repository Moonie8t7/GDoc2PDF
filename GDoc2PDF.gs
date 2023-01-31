/**
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @file https://www.reddit.com/r/GoogleAppsScript/comments/10p0mfm/automatic_conversion_of_gdocs_to_pdfs
 * @desc Takes a parent folder, then iterates through it and its children and converts each Google Doc to a PDF in a new folder.
 * @license MIT
 * @version 1.1
 */

/* It's the ID of the folder you want to convert the files in. */
const FOLDER_ID = "your top level folder ID goes here";

/* It's the number of files that will be converted at a time. */
const BATCH_SIZE = 50;

/**
 * Gets all the files in the top folder, then it gets all the files in all the subfolders, then it
 * converts all the files in the top folder and all the subfolders to PDF
 * @throws {Error} If an error occurs while getting the folder or converting the files to PDF
 * @example
 * convertToPDF();
 */
function convertToPDF() {
  try {
    const topFolder = DriveApp.getFolderById(FOLDER_ID);
    convertFilesInFolder_(topFolder);
  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}

/**
 * Takes a folder, checks if it's name contains "CONVERTED", if not, it gets all the Google Docs in
 * the folder, creates a new folder with the name "CONVERTED" appended to the original folder name,
 * then converts the Google Docs to PDFs and puts them in the new folder
 * @param folder - The folder you want to convert the files in.
 * @throws {Error} If there is an error during processing.
 * @example
 * convertFilesInFolder_(DriveApp.getFolderById(FOLDER_ID))
 */
function convertFilesInFolder_(folder) {
  try {
    if (!folder.getName().includes("CONVERTED")) {
      const docs = folder.getFilesByType(MimeType.GOOGLE_DOCS);
      const docArray = [];
      while (docs.hasNext()) {
        docArray.push(docs.next());
      }
      const convertedFolder = createConvertedFolder_(folder);
      for (let i = 0; i < docArray.length; i += BATCH_SIZE) {
        const batch = docArray.slice(i, i + BATCH_SIZE);
        Utilities.sleep(1000);
        batch.map((doc) =>
          createPDFfile_(doc.getId(), convertedFolder.getId())
        );
      }
    }
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      convertFilesInFolder_(subFolder);
    }
  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}

/**
 * Creates a folder named "folderName_CONVERTED" in the folder passed to the function.
 * If the folder already exists, it returns the existing folder.
 * @param folder - The folder you want to convert.
 * @returns the converted folder.
 * @throws {Error} If there's an error while creating the converted folder.
 * @example
 * const folder = DriveApp.getFolderById("folderId");
 * const convertedFolder = createConvertedFolder_(folder);
 */
function createConvertedFolder_(folder) {
  try {
    const folderName = `${getFileName_(folder.getName())}_CONVERTED`;
    let convertedFolder;
    const existingFolders = folder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      convertedFolder = existingFolders.next();
    } else {
      convertedFolder = folder.createFolder(folderName);
    }
    return convertedFolder;
  } catch (error) {
    console.error(`Error while creating converted folder: ${error.message}`);
  }
}

/**
 * Takes a file ID and a folder ID as parameters, and if the file is not already in the folder,
 * create a PDF version of the file and move it to the folder. It also checks if the file is newer than
 * the PDF file in the folder, it creates a new PDF file in the folder replacing the old PDF.
 * @param fileID - The ID of the file you want to convert to PDF.
 * @param folderID - The ID of the folder where you want to save the PDF file.
 * @throws Will throw an error if there is an issue with converting the file to PDF or moving it to the folder.
 * @example
 * createPDFfile_("abc123", "def456");
 */
function createPDFfile_(fileID, folderID) {
  try {
    const templateFile = DriveApp.getFileById(fileID);
    const folder = DriveApp.getFolderById(folderID);
    const fileName = `${getFileName_(templateFile.getName())}.pdf`;
    let existingFile = folder.getFilesByName(fileName).hasNext()
      ? folder.getFilesByName(fileName).next()
      : null;
    if (
      existingFile &&
      existingFile.getDateCreated().valueOf() >=
        templateFile.getLastUpdated().valueOf()
    ) {
      return;
    }
    if (existingFile) {
      existingFile.setTrashed(true);
    }
    const theBlob = templateFile.getBlob().getAs("application/pdf");
    folder.createFile(theBlob).setName(fileName);
    console.log(`Success: PDF created for file ${fileName}`);
  } catch (error) {
    console.error(
      `Error while converting to PDF and moving to folder: ${error.message}`
    );
  }
}

/**
 * Takes a string, finds the last period in the string, and returns the string without the last
 * period and everything after it.
 * @param fileName - The name of the file you want to get the name of.
 * @returns The file name without the extension.
 * @example
 * getFileName_("myFile.txt") // returns "myFile"
 */
function getFileName_(fileName) {
  return fileName.replace(/\.[^/.]+$/, "");
}
