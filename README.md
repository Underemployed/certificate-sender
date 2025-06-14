# Certificate Automation Script

This Google Apps Script automates the creation and distribution of certificates using Google Slides, Sheets, and Gmail. Designed for event organizers to efficiently manage certificate generation and email delivery.

## Features
- 📄 **Auto-generate certificates** from a Google Slides template
- 📧 **Send certificates via email** as PDF attachments
- ✅ **Status tracking** in Google Sheets
- 🔄 **Placeholder replacement** for participant details
- 🛠️ **Automatic sheet configuration**

## Setup Instructions

### 1. Use Template
1. Access the template: [Drive Template](https://drive.google.com/drive/folders/1VXmOnYeCrbmjNWG8g1RNoNK9diYslLCJ?usp=sharing)
2. Make copies of all required files

### 2. Configuration
Update these variables in the script:
```javascript
var eventName = "Master the Basics of Flutter";
var SocietyName = "ISTE SC GECBH";
var slideTemplateUrl = "YOUR_SLIDE_TEMPLATE_URL";
var tempFolderUrl = "YOUR_TEMP_FOLDER_URL";
var sheetUrl = "YOUR_SHEET_URL";
```

### 3. Google Sheet Preparation
Create a Google Sheet with these columns (order doesn't matter):
- **Name** - Full name of the participant 
- **Email** - Email address of the participant
- **College** - Institution name of the participant
- **Slide ID** (auto-populated)
- **Status** (auto-populated)

### 4. Template Presentation Setup

Ensure the template has these placeholders:
- `<NAME>` for participant names
- `<COLLEGE>` for institution names

## Usage

### 1. Create Certificates
1. Populate participant data in the appscript
2. Run `createCertificates()` from the script editor. You can see live changes in the sheet.
3. Monitor status column for progress

### 2. Send Certificates
1. Verify certificates marked as "CREATED" in sheet.
2. Run `sendCertificates()` from the script editor
3. Check status for "SENT" confirmation.

## Error Handling
Monitor "Status" column for:
- `Missing data` - Incomplete information
- `Missing email/slide` - Missing required fields
- `ERROR: [message]` - Specific error details. Usually timeout just rerun.

## Important Notes
1. Template must contain `<NAME>` and `<COLLEGE>` placeholders
2. Certificates are stored in specified temp folder
3. Runs in batches of 40 to avoid timeout issues

