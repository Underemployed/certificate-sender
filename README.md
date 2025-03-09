# Certificate Automation Script

This Google Apps Script automates the creation and distribution of certificates using Google Slides, Sheets, and Gmail. Designed for event organizers to efficiently manage certificate generation and email delivery.

## Features
- 📄 **Auto-generate certificates** from a Google Slides template
- 📧 **Send certificates via email** as PDF attachments
- ✅ **Status tracking** in Google Sheets
- 🔄 **Placeholder replacement** for participant details
- 🛠️ **Automatic sheet configuration**

## Setup Instructions

### 1. Configuration
Update these variables in the script:
```javascript
var eventName = "Gitflow 2.0"; // Your event name
var SocietyName = "ISTE SC GECBH"; // Your organization name
var slideTemplateUrl = "YOUR_SLIDE_TEMPLATE_URL"; // Google Slides URL
var tempFolderUrl = "YOUR_TEMP_FOLDER_URL"; // Google Drive folder URL
var sheetUrl = "YOUR_SHEET_URL"; // Google Sheets URL
```

### 2. Google Sheet Preparation
Create a Google Sheet with these columns (order doesn't matter):
- **Name** - Participant's full name
- **Email** - Participant's email address
- **College** - Participant's institution
- **Slide ID** (auto-populated)
- **Status** (auto-populated)


### 3. Template Setup
Create a Google Slides template with these placeholders:
- `<NAME>` for participant names
- `<COLLEGE>` for institution names

## Usage

### 1. Create Certificates
1. Populate participant data in the sheet
2. Run `createCertificates()` from the script editor
3. Certificates will be:
   - Generated in your temp folder
   - Linked in the "Slide ID" column
   - Marked as "CREATED" when successful

### 2. Send Certificates
1. Ensure all certificates are marked "CREATED"
2. Run `sendCertificates()` from the script editor
3. Emails will be:
   - Sent with PDF attachments
   - Marked as "SENT" when successful


## Error Handling
Check the "Status" column for:
- `Missing data` - Incomplete participant information
- `Missing email/slide` - Required fields for emailing
- `ERROR: [message]` - Specific error details

## Important Notes
1. Template Requirements:
   - Must contain `<NAME>` and `<COLLEGE>` placeholders
   - All slides will be processed for text replacement
   
2. Folder Structure:
   ```plaintext
   📁 Your Temp Folder
   └── 📄 [Participant Name] - Certificate (Google Slide)
   ```




