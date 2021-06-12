/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
 */
const RECIPIENT_COL = "Student Email";
const EMAIL_SENT_COL = "Status";

var slideTemplateId = "1Mh0LP9pSRvREDPzrB_d_PEPxwi0iYZRaV5n7tjUgf3Y"; // Sample: https://docs.google.com/spreadsheets/d/1cgK1UETpMF5HWaXfRE6c0iphWHhl7v-dQ81ikFtkIVk
var tempFolderId = "1nxR1eH5U3JDMcQOHXSPqGOWdO24GlXme"; // Create an empty folder in Google Drive

var CREATED = 'CREATED';
var EMAILED = 'EMAILED';

/**
 * Creates a custom menu "Diploma" in the spreadsheet
 * with drop-down options to create and send certificates
 */
function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Diploma')
        .addItem('Create certificates', 'createCertificates')
        .addSeparator()
        .addItem('Send certificates', 'sendCertificates2')
        .addToUi();
}

/**
 * Creates a personalized certificate for each student
 * and stores every individual Slides doc on Google Drive
 */
function createCertificates() {

    // Load the Google Slide template file
    var template = DriveApp.getFileById(slideTemplateId);

    // Get all student data from the spreadsheet and identify the headers
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var empNameIndex = headers.indexOf("Student Name");
    //var empEmailIndex = headers.indexOf("Student Email");
    var empSlideIndex = headers.indexOf("Student Slide");
    var statusIndex = headers.indexOf("Status");
    var empLinkIndex = headers.indexOf("Student Slide Link");

    // Iterate through each row to capture individual details
    for (var i = 1; i < values.length; i++) {
        var rowData = values[i];
        var empName = rowData[empNameIndex];
        var statusUpdate = rowData[statusIndex];

        if (statusUpdate !== CREATED && statusUpdate !== EMAILED) { //Checks if a certificate hasn't already been made AND sent
            // Make a copy of the Slide template and rename it with student name
            var tempFolder = DriveApp.getFolderById(tempFolderId);
            var empSlideId = template.makeCopy(tempFolder).setName(empName).getId();
            var empSlide = SlidesApp.openById(empSlideId).getSlides()[0];


            // Replace placeholder values with actual student related details
            empSlide.replaceAllText("Student Name", empName);

            // Update the spreadsheet with the new Slide Id and status
            sheet.getRange(i + 1, empSlideIndex + 1).setValue(empSlideId);
            sheet.getRange(i + 1, statusIndex + 1).setValue("CREATED");
            sheet.getRange(i + 1, empLinkIndex + 1).setValue("https://docs.google.com/presentation/d/" + empSlideId);
            SpreadsheetApp.flush();
        }
    }
}

/**
 * Send an email to each individual student
 * with a PDF attachment of their appreciation certificate
 */
function sendCertificates() {

    // Get all student data from the spreadsheet and identify the headers
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var empNameIndex = headers.indexOf("Student Name");
    var empEmailIndex = headers.indexOf("Student Email");
    var empSlideIndex = headers.indexOf("Student Slide");
    var statusIndex = headers.indexOf("Status");

    // Iterate through each row to capture individual details
    for (var i = 1; i < values.length; i++) {
        var rowData = values[i];
        var empName = rowData[empNameIndex];
        var empSlideId = rowData[empSlideIndex];
        var empEmail = rowData[empEmailIndex];
        var statusUpdate = rowData[statusIndex];

        if (statusUpdate !== EMAILED) { //Checks if an certificate hasn't already been sent

            // Load the student's personalized Google Slide file
            var attachment = DriveApp.getFileById(empSlideId);

            // Setup the required parameters and send them the email
            var senderName = "URSP Advisory Board";
            var subject = empName + ", congrats on your URSP invitation!";
            var body = "Congratulations on your invitation to the incoming 2021 University Research Scholars Program (URSP) Cohort!" + "\n" + "You are now part of a group of top students who will be given a comprehensive set of opportunities to help you make the most of your time at UF and prepare you for any path you choose to follow. You will be receiving a survey in the coming days that you must complete before May 1 in order to confirm your acceptance of this invitation and join URSP. We look forward to welcoming you when you arrive on campus."
    + "\n\n" + "Go URSP! Go Gators! Go Research!" + "\n" + "Dr. Donnelly" + "\n" + "Director, UF Center for Undergraduate Research" + "\n\n" + "Note: if there is a mistake on your personalized attachment, please feel free to reply to this email and we can make any necessary corrections.";

            GmailApp.sendEmail(empEmail, subject, body, {
                attachments: [attachment.getAs(MimeType.PDF)],
                name: senderName
            });

            // Update the spreadsheet with email status
            sheet.getRange(i + 1, statusIndex + 1).setValue("EMAILED");
            SpreadsheetApp.flush();
        }
    }
}

/**
 * Send an email to each individual student
 * with a PDF attachment of their appreciation certificate
 */
function sendCertificates2() {

    // Get all student data from the spreadsheet and identify the headers
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var empNameIndex = headers.indexOf("Student Name");
    var empEmailIndex = headers.indexOf("Student Email");
    var empSlideIndex = headers.indexOf("Student Slide");
    var statusIndex = headers.indexOf("Status");

    // option to skip browser prompt if you want to use this code in other projects  
    /*var subjectLine = Browser.inputBox("Mail Merge", 
                                        "Type or copy/paste the subject line of the Gmail " +
                                        "draft message you would like to mail merge with:",
                                        Browser.Buttons.OK_CANCEL);
                                        
      if (subjectLine === "cancel" || subjectLine == ""){ 
      // if no subject line finish up
      return;
      }*/

    var subjectLine = 'Congrats on your URSP admission!';

    // get the draft Gmail message to use as a template
    const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

    // Iterate through each row to capture individual details
    for (var i = 1; i < values.length; i++) {
        var rowData = values[i];
        var empName = rowData[empNameIndex];
        var empSlideId = rowData[empSlideIndex];
        var empEmail = rowData[empEmailIndex];
        var statusUpdate = rowData[statusIndex];

        if (statusUpdate !== EMAILED) { //Checks if an certificate hasn't already been sent

            const msgObj = fillInTemplateFromObject_(emailTemplate.message, rowData);

            // Load the student's personalized Google Slide file
            var attachment = DriveApp.getFileById(empSlideId);

            // Setup the required parameters and send them the email
            var senderName = "URSP Advisory Board";
            var subject = empName + ", congrats on your URSP admission!";

            // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
            // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
            // Uncomment advanced parameters as needed (see docs for limitations)
            GmailApp.sendEmail(empEmail, subject, msgObj.text, {
                attachments: [attachment.getAs(MimeType.PDF)],
                name: senderName,
                inlineImages: emailTemplate.inlineImages
            });

            // Update the spreadsheet with email status
            sheet.getRange(i + 1, statusIndex + 1).setValue("EMAILED");
            SpreadsheetApp.flush();
        }
    }
}

/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */
function getGmailTemplateFromDrafts_(subject_line) {
    try {
        // get drafts
        const drafts = GmailApp.getDrafts();
        // filter the drafts that match subject line
        const draft = drafts.filter(subjectFilter_(subject_line))[0];
        // get the message object
        const msg = draft.getMessage();

        // Handling inline images and attachments so they can be included in the merge
        // Based on https://stackoverflow.com/a/65813881/1027723
        // Get all attachments and inline image attachments
        const allInlineImages = draft.getMessage().getAttachments({
            includeInlineImages: true,
            includeAttachments: false
        });
        const attachments = draft.getMessage().getAttachments({
            includeInlineImages: false
        });
        const htmlBody = msg.getBody();

        // Create an inline image object with the image name as key 
        // (can't rely on image index as array based on insert order)
        const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

        //Regexp to search for all img string positions with cid
        const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
        const matches = [...htmlBody.matchAll(imgexp)];

        //Initiate the allInlineImages object
        const inlineImagesObj = {};
        // built an inlineImagesObj from inline image matches
        matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

        return {
            message: {
                subject: subject_line,
                text: msg.getPlainBody(),
                html: htmlBody
            },
            attachments: attachments,
            inlineImages: inlineImagesObj
        };
    } catch (e) {
        throw new Error("Oops - can't find Gmail draft");
    }
}

/**
 * Filter draft objects with the matching subject linemessage by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} GmailDraft object
 */
function subjectFilter_(subject_line) {
    return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
            return element;
        }
    }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
        return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
    return str
        .replace(/[\\]/g, '\\\\')
        .replace(/[\"]/g, '\\\"')
        .replace(/[\/]/g, '\\/')
        .replace(/[\b]/g, '\\b')
        .replace(/[\f]/g, '\\f')
        .replace(/[\n]/g, '\\n')
        .replace(/[\r]/g, '\\r')
        .replace(/[\t]/g, '\\t');
}