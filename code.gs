/*********************** CONFIGURATION ***********************/
var eventName = "Gitflow 2.0";
var slideTemplateUrl = "https://docs.google.com/presentation/d/1HqRF0cAo_1PmYszSqQZjmYkA1Z2C_RTO3aNlNXNaCC0/edit";
var tempFolderUrl = "https://drive.google.com/drive/folders/1cYzG-Hj_jG2uFoyffn47yeHcfzubbzwP";
var SocietyName = "ISTE SC GECBH";
var sheetUrl = "https://docs.google.com/spreadsheets/d/1MvTSRpp0yfpeYfmajukfsjelLEct1G0F_-2Po5ebIs0/edit";

/*********************** HELPER FUNCTIONS ***********************/
function getIdFromUrl(url) {
    const pattern = /[-\w]{25,}/;
    const match = url.match(pattern);
    return match ? match[0] : null;
}

function selectOrCreateSheet(sheetName) {
    const ss = SpreadsheetApp.openByUrl(sheetUrl);
    let sheet = ss.getSheetByName(sheetName);
    return sheet ? sheet : ss.insertSheet(sheetName);
}

function getColumnIndex(sheet, column) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.findIndex(h => h.toLowerCase() === column.toLowerCase());
}

/*********************** LOGGING FUNCTIONS ***********************/
function log(message) {
    const logSheet = selectOrCreateSheet("Process Logs");
    logSheet.appendRow([new Date(), message]);
    Logger.log(message);
}

function logError(error) {
    const errorSheet = selectOrCreateSheet("Error Logs");
    errorSheet.appendRow([
        new Date(),
        error.message,
        error.stack || "No stack trace available"
    ]);
    Logger.severe(`ERROR: ${error.message}\n${error.stack}`);
}

/*********************** CORE FUNCTIONS ***********************/
function setupSheet() {
    try {
        const sheet = selectOrCreateSheet("Sheet1");
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Add missing columns
        ['Slide ID', 'Status'].forEach(col => {
            if (!headers.some(h => h.toLowerCase() === col.toLowerCase())) {
                sheet.getRange(1, headers.length + 1).setValue(col);
                headers.push(col);
            }
        });

        // Verify required columns
        const requiredColumns = ['Name', 'Email', 'College'];
        requiredColumns.forEach(col => {
            if (!headers.some(h => h.toLowerCase() === col.toLowerCase())) {
                throw new Error(`Missing required column: ${col}`);
            }
        });

        return {
            sheet: sheet,
            nameIndex: getColumnIndex(sheet, "Name"),
            emailIndex: getColumnIndex(sheet, "Email"),
            collegeIndex: getColumnIndex(sheet, "College"),
            slideIndex: getColumnIndex(sheet, "Slide ID"),
            statusIndex: getColumnIndex(sheet, "Status")
        };
    } catch (error) {
        log("Setup failed");
        throw error;
    }
}

function createCertificates(setup) {
    const { sheet } = setup;
    const data = sheet.getDataRange().getValues().slice(1);

    try {
        const template = DriveApp.getFileById(getIdFromUrl(slideTemplateUrl));
        const folder = DriveApp.getFolderById(getIdFromUrl(tempFolderUrl));

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowNumber = i + 2;
            const status = row[setup.statusIndex]?.toUpperCase();

            if (['CREATED', 'SENT'].includes(status)) continue;

            const [name, college] = [row[setup.nameIndex], row[setup.collegeIndex]];
            if (!name || !college) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue('Missing data');
                throw new Error(`Row ${rowNumber}: Missing name or college`);
            }

            try {
                const slideCopy = template.makeCopy(`${name} - Certificate`, folder);
                const presentation = SlidesApp.openById(slideCopy.getId());

                presentation.getSlides().forEach(slide => {
                    slide.replaceAllText("<NAME>", name);
                    slide.replaceAllText("<COLLEGE>", college);
                });

                sheet.getRange(rowNumber, setup.slideIndex + 1).setValue(slideCopy.getId());
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("CREATED");
                log(`Created slide for ${name}`);
            } catch (error) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
                throw error;
            }
        }
    } catch (error) {
        log("Certificate creation failed");
        throw error;
    }
}

function sendCertificates(setup) {
    const { sheet } = setup;
    const data = sheet.getDataRange().getValues().slice(1);

    try {
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowNumber = i + 2;
            const status = row[setup.statusIndex]?.toUpperCase();

            if (status !== 'CREATED') continue;

            const [name, email, slideId] = [
                row[setup.nameIndex],
                row[setup.emailIndex],
                row[setup.slideIndex]
            ];

            if (!email || !slideId) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue('Missing email/slide');
                throw new Error(`Row ${rowNumber}: Missing email or slide ID`);
            }

            try {
                const pdfFile = DriveApp.getFileById(slideId).getAs(MimeType.PDF);
                const subject = `${name}, Your ${eventName} Certificate`;
                const body = `Dear ${name},\n\nThank you for participating in ${eventName} organized by ${SocietyName}.\n\nFind your certificate attached.\n\nBest regards,\n${SocietyName}`;

                GmailApp.sendEmail(email, subject, body, {
                    attachments: [pdfFile],
                    name: SocietyName
                });

                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("SENT");
                log(`Sent certificate to ${email}`);
            } catch (error) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
                throw error;
            }
        }
    } catch (error) {
        log("Certificate sending failed");
        throw error;
    }
}

/*********************** MAIN FUNCTION ***********************/
function main() {
    try {
        log("=== Starting Certificate Generation Process ===");

        log("Setting up spreadsheet...");
        const setup = setupSheet();
        log("Spreadsheet setup completed successfully");

        log("Starting certificate creation...");
        createCertificates(setup);
        log("Certificate creation completed successfully");

        log("Starting certificate distribution...");
        sendCertificates(setup);
        log("Certificate distribution completed successfully");

        log("=== Process Completed Successfully ===");
    } catch (error) {
        log("Process aborted due to error");
        logError(error);
        throw error; 
    }
}