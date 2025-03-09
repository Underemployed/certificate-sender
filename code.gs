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

/*********************** CORE FUNCTIONS ***********************/
function setupSheet() {
    const sheet = selectOrCreateSheet("Sheet1");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Add missing columns if needed
    ['Slide ID', 'Status'].forEach(col => {
        if (!headers.some(h => h.toLowerCase() === col.toLowerCase())) {
            sheet.getRange(1, headers.length + 1).setValue(col);
            headers.push(col);
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
}

function createCertificates() {
    const setup = setupSheet();
    const { sheet } = setup;

    // Validate required columns
    if ([setup.nameIndex, setup.emailIndex, setup.collegeIndex].some(i => i === -1)) {
        throw new Error('Missing required columns: Name, Email, or College');
    }

    const template = DriveApp.getFileById(getIdFromUrl(slideTemplateUrl));
    const folder = DriveApp.getFolderById(getIdFromUrl(tempFolderUrl));
    const data = sheet.getDataRange().getValues();

    data.slice(1).forEach((row, index) => {
        const rowNumber = index + 2;
        const status = row[setup.statusIndex]?.toUpperCase();
        if (['CREATED', 'SENT'].includes(status)) return;

        try {
            const [name, college] = [row[setup.nameIndex], row[setup.collegeIndex]];
            if (!name || !college) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue('Missing data');
                return;
            }

            // Create certificate
            const slideCopy = template.makeCopy(`${name} - Certificate`, folder);
            const presentation = SlidesApp.openById(slideCopy.getId());

            presentation.getSlides().forEach(slide => {
                slide.replaceAllText("<NAME>", name);
                slide.replaceAllText("<COLLEGE>", college);
            });

            // Update sheet
            sheet.getRange(rowNumber, setup.slideIndex + 1).setValue(slideCopy.getId());
            sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("CREATED");
        } catch (error) {
            sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
        }
    });
}

function sendCertificates() {
    const setup = setupSheet();
    const { sheet } = setup;
    const data = sheet.getDataRange().getValues();

    data.slice(1).forEach((row, index) => {
        const rowNumber = index + 2;
        const status = row[setup.statusIndex]?.toUpperCase();
        if (status !== 'CREATED') return;

        try {
            const [name, email, slideId] = [
                row[setup.nameIndex],
                row[setup.emailIndex],
                row[setup.slideIndex]
            ];

            if (!email || !slideId) {
                sheet.getRange(rowNumber, setup.statusIndex + 1).setValue('Missing email/slide');
                return;
            }

            // Prepare email
            var pdfFile = DriveApp.getFileById(slideId).getAs(MimeType.PDF);
            var subject = `${name}, Your ${eventName} Certificate`;

            var body = `Dear ${name},\n\n` +
                `Thank you for participating in ${eventName} organized by ${SocietyName}\n\n` +
                `Find your certificate attached. We appreciate your participation.\n\n` +
                `Best regards,\n${SocietyName}`;

            // Send email
            GmailApp.sendEmail(email, subject, body, {
                attachments: [pdfFile],
                name: SocietyName
            });

            sheet.getRange(rowNumber, setup.statusIndex + 1).setValue("SENT");
        } catch (error) {
            sheet.getRange(rowNumber, setup.statusIndex + 1).setValue(`ERROR: ${error.message}`);
        }
    });
}