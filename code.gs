/*********************** CONFIGURATION ***********************/
var eventName = "Gitflow 2.0";
var SocietyName = "ISTE SC GECBH";
var slideTemplateUrl = "https://docs.google.com/presentation/d/1HqRF0cAo_1PmYszSqQZjmYkA1Z2C_RTO3aNlNXNaCC0/edit";
var tempFolderUrl = "https://drive.google.com/drive/folders/1cYzG-Hj_jG2uFoyffn47yeHcfzubbzwP";
var sheetUrl = "https://docs.google.com/spreadsheets/d/1MvTSRpp0yfpeYfmajukfsjelLEct1G0F_-2Po5ebIs0/edit";

/*********************** HELPER FUNCTIONS ***********************/
function getIdFromUrl(url) {
    const pattern = /[-\w]{25,}/;
    const match = url.match(pattern);
    return match ? match[0] : null;
}

function selectOrCreateSheet(sheetName) {
    try {
        const ss = SpreadsheetApp.openByUrl(sheetUrl);
        let sheet = ss.getSheetByName(sheetName);
        return sheet ? sheet : ss.insertSheet(sheetName);
    } catch (error) {
        Logger.log(`Error in selectOrCreateSheet: ${error.message}`);
        throw error;
    }
}

function getColumnIndex(sheet, column) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const index = headers.findIndex(h => h.toLowerCase() === column.toLowerCase());
    if (index === -1) throw new Error(`Column '${column}' not found`);
    return index;
}

/*********************** CORE FUNCTIONS ***********************/
function setupSheet() {
    try {
        Logger.log("Starting sheet setup...");
        const sheet = selectOrCreateSheet("Sheet1");
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Add missing columns if needed
        ['Slide ID', 'Status'].forEach(col => {
            if (!headers.some(h => h.toLowerCase() === col.toLowerCase())) {
                Logger.log(`Adding missing column: ${col}`);
                sheet.getRange(1, headers.length + 1).setValue(col);
                headers.push(col);
            }
        });

        const indices = {
            sheet: sheet,
            nameIndex: getColumnIndex(sheet, "Name"),
            emailIndex: getColumnIndex(sheet, "Email"),
            collegeIndex: getColumnIndex(sheet, "College"),
            slideIndex: getColumnIndex(sheet, "Slide ID"),
            statusIndex: getColumnIndex(sheet, "Status")
        };

        Logger.log("Sheet setup completed successfully");
        return indices;
    } catch (error) {
        Logger.log(`Sheet setup failed: ${error.message}`);
        throw error;
    }
}

function createCertificates() {
    try {
        Logger.log("Starting certificate creation...");
        const setup = setupSheet();
        const { sheet } = setup;

        const template = DriveApp.getFileById(getIdFromUrl(slideTemplateUrl));
        const folder = DriveApp.getFolderById(getIdFromUrl(tempFolderUrl));
        const data = sheet.getDataRange().getValues();

        Logger.log(`Processing ${data.length - 1} participants`);

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const rowNumber = i + 1;
            const status = row[setup.statusIndex]?.toUpperCase();

            try {
                if (['CREATED', 'SENT'].includes(status)) {
                    Logger.log(`Skipping row ${rowNumber} - Status: ${status}`);
                    continue;
                }

                const [name, college] = [row[setup.nameIndex], row[setup.collegeIndex]];
                if (!name || !college) {
                    const errorMsg = `Missing data in row ${rowNumber}`;
                    sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(errorMsg);
                    throw new Error(errorMsg);
                }
                Logger.log(`Creating certificate for ${name} (Row ${rowNumber})`);

                const slideCopy = template.makeCopy(`${name} - Certificate`, folder);
                const presentation = SlidesApp.openById(slideCopy.getId());


                presentation.getSlides().forEach(slide => {
                    slide.replaceAllText("<NAME>", name);
                    slide.replaceAllText("<COLLEGE>", college);
                });

                sheet.getRange(rowNumber, setup.slideIndex + 1).setValue(slideCopy.getId());
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("CREATED");
                Logger.log(`created certificate for ${name} - ${slideCopy.getUrl()}`);
            } catch (error) {
                Logger.log(`Error in row ${rowNumber}: ${error.message}`);
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
                throw error; // Stop execution on first error
            }
        }
        Logger.log("Certificate creation process completed");
    } catch (error) {
        Logger.log(`Certificate creation failed: ${error.message}`);
        throw error;
    }
}

function sendCertificates() {
    try {
        Logger.log("Starting certificate distribution...");
        const setup = setupSheet();
        const { sheet } = setup;
        const data = sheet.getDataRange().getValues();

        Logger.log(`Checking ${data.length - 1} participants for sending`);

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const rowNumber = i + 1;
            const status = row[setup.statusIndex]?.toUpperCase();

            try {
                if (status !== 'CREATED') {
                    Logger.log(`Skipping row ${rowNumber} - Status: ${status}`);
                    continue;
                }

                const [name, email, slideId] = [
                    row[setup.nameIndex],
                    row[setup.emailIndex],
                    row[setup.slideIndex]
                ];

                if (!email || !slideId) {
                    const errorMsg = `Missing email/slide in row ${rowNumber}`;
                    sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(errorMsg);
                    throw new Error(errorMsg);
                }

                Logger.log(`Sending certificate to ${email} (Row ${rowNumber})`);

                const pdfFile = DriveApp.getFileById(slideId).getAs(MimeType.PDF);
                const subject = `${name}, Your ${eventName} Certificate`;
                const body = `Dear ${name},\n\nThank you for participating in ${eventName} organized by ${SocietyName}\n\nFind your certificate attached.\n\nBest regards,\n${SocietyName}`;

                GmailApp.sendEmail(email, subject, body, {
                    attachments: [pdfFile],
                    name: SocietyName
                });

                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("SENT");
                Logger.log(`Successfully sent to ${email}`);
            } catch (error) {
                Logger.log(`Error in row ${rowNumber}: ${error.message}`);
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
                throw error; // Stop execution on first error
            }
        }
        Logger.log("Certificate distribution completed");
    } catch (error) {
        Logger.log(`Certificate distribution failed: ${error.message}`);
        throw error;
    }
}

/*********************** MAIN FUNCTION ***********************/
function main() {
    try {
        TEST = false; // dont send mail
        Logger.log("Starting certificate automation process");

        // Create certificates
        createCertificates();
        Logger.log("Certificates created successfully");


        if (!TEST) {
            Logger.log("User confirmed to send certificates");
            sendCertificates();
        }

        Logger.log("Process completed successfully");
    } catch (error) {
        Logger.log(`Process failed: ${error.message}`);

        throw error;
    }
}