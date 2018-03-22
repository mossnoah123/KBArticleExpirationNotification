// Author: Noah Moss
// Date Created: 2/26/2018
// Date Updated: 3/13/2018
// Purpose: This script will automatically send the TTC an email whenever it detects that a KB article is about to expire.

// Constants
var EMAIL_RECIPIENT = "ttc@uwplatt.edu";
var EMAIL_SUBJECT = "Attn: KB Articles expiring soon";
var SHEET_COLUMN_RANGE = "A:G";
var ARTICLE_SHEET_NAME = "KBArticles"
var START_ARTICLE_ROW = 2;
var TOTAL_COLUMNS = 7;
var DAYS_IN_WEEK = 7;

// Article Struct
var Article = function (id, title, status, owner, expiration, dateChecked, checkedBy) {
    this.id = id;
    this.title = title;
    this.status = status;
    this.owner = owner;
    this.expiration = expiration;
    this.dateChecked = dateChecked;
    this.checkedBy = checkedBy;
}

// Email body build Enumerated Type
var BuildEmailBodyTypes = {
    "New":0,
    "Insert":1,
    "Conclude":2
}
Object.freeze(BuildEmailBodyTypes);


/**
 * Runs the main script.
 */
function main() {
    var currentDate = new Date();
    var kbSheet = getKBArticleSheet();
    var kbArticlesInSheet = getKBArticlesWithinSheet(kbSheet);
    var emailBody = buildEmailBody(BuildEmailBodyTypes.New, null, currentDate);
    for each(var kbArticle in kbArticlesInSheet)
    {
        if (isArticleExpiring(kbArticle, currentDate))
            emailBody += buildEmailBody(BuildEmailBodyTypes.Insert, kbArticle, null);
    }
    if (articlesWereInsertedIntoEmailBody(emailBody, currentDate)) {
        emailBody += buildEmailBody(BuildEmailBodyTypes.Conclude, null, null);
        sendEmail(emailBody);
    }
}

/**
 * Returns the sheet that the list of KB Articles is on.
 */
function getKBArticleSheet() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARTICLE_SHEET_NAME);
}

/**
 * Gets the data contents of every cell from each KB Article row, constructs a new KB Article
 * object using those data contents, adds that new object to an array of KB Article objects, 
 * and finally returns the array.
 * 
 * @param {Sheet} sheet - Sheet to get KB Articles from.
 */
function getKBArticlesWithinSheet(sheet) {
    var rawKBArticleData = sheet.getRange(SHEET_COLUMN_RANGE).getValues();
    var arrayOfKBArticleObjects = new Array();
    for (var i = START_ARTICLE_ROW - 1, l = rawKBArticleData.length; i < l; i++) {
        if (rawKBArticleData[i][0].toString() == "") // ID of record
            break;
        else {
            var article = constructNewKBArticle(rawKBArticleData, i);
            arrayOfKBArticleObjects.push(article);
        }
    }
    return arrayOfKBArticleObjects;
}

/**
 * Constructs and returns a new KB Article object.
 * 
 * @param {Object[][]} kbArticleData - Container holding the raw data of every KB Article.
 * @param {Integer} rowIndex - Index of the KB Article row to use.
 */
function constructNewKBArticle(kbArticleData, rowIndex) {
    var id = kbArticleData[rowIndex][0];
    var title = kbArticleData[rowIndex][1];
    var status = kbArticleData[rowIndex][2];
    var owner = kbArticleData[rowIndex][3];
    var expiration = kbArticleData[rowIndex][4];
    var dateChecked = kbArticleData[rowIndex][5];
    var checkedBy = kbArticleData[rowIndex][6];
    var kbArticle = new Article(id, title, status, owner, expiration, dateChecked, checkedBy);
    return kbArticle;
}

/**
 * Uses the current date to determine whether a specified article is nearing it's
 * expiration date.
 * 
 * @param {Article} article - The KB Article used to determine if it's close to expiration.
 * @param {Date} currentDate - Today's date.
 */
function isArticleExpiring(article, currentDate) {
    var weekFromExpirationDate = getPreviousWeekDate(article.expiration);
    if (isCurrentDateOneWeekFromExpiring(currentDate, weekFromExpirationDate))
        return true;
    return false;
}

/**
 * Returns the date of the previous week from the specified date.
 * 
 * @param {Date} date - Date to get the previous week from.
 */
function getPreviousWeekDate(date) {
    var weekPriorDate = new Date(date);
    var dayOfMonth = weekPriorDate.getDate();
    weekPriorDate.setDate(dayOfMonth - DAYS_IN_WEEK);
    return weekPriorDate;
}

/**
 * Returns the date of the next week from the specified date.
 * 
 * @param {Date} date - Date to get the next week from.
 */
function getNextWeekDate(date) {
    var weekNextDate = new Date(date);
    var dayOfMonth = weekNextDate.getDate();
    weekNextDate.setDate(dayOfMonth + DAYS_IN_WEEK);
    return weekNextDate;
}

/**
 * Determines whether the current date is equal to the date that is one week away
 * from an article's expiration date.
 * 
 * @param {Date} currentDate - Today's date.
 * @param {Date} weekFromExpirationDate - The date one week prior to an article's expiration date.
 */
function isCurrentDateOneWeekFromExpiring(currentDate, weekFromExpirationDate) {
    var currentDateString = getDateInStringFormat(currentDate);
    var weekFromExpirationDateString = getDateInStringFormat(weekFromExpirationDate);
    if (currentDateString == weekFromExpirationDateString)
        return true;
    return false;
}

/**
 * Returns the "mm/dd/yyyy" format of a specified date.
 * 
 * @param {Date} date - Date to get the "mm/dd/yyyy" format for.
 */
function getDateInStringFormat(date) {
    var dateDay = date.getDate();
    var dateMonth = date.getMonth();
    var dateYear = date.getFullYear();
    return (dateMonth + 1) + "/" + dateDay + "/" + dateYear;
}

/**
 * Determines whether any articles were inserted into the body of the email to be sent.
 * 
 * @param {String} emailBody - Body of the email.
 * @param {Date} currentDate - Today's date.
 */
function articlesWereInsertedIntoEmailBody(emailBody, currentDate) {
    if (emailBody == newEmailBody(currentDate))
        return false;
    return true;
}

/**
 * Uses Google's mail sending service to send the fully composed email detailing all of
 * the KB Articles that are about to expire.
 * 
 * @param {String} emailBody - Body of the email.
 */
function sendEmail(emailBody) {
    MailApp.sendEmail(EMAIL_RECIPIENT, EMAIL_SUBJECT, emailBody);
}

/**
 * Builds and returns a portion of the email body, depending on the specified build type.
 * 
 * @param {BuildEmailBodyTypes} buildType - Specifies which part of the email body to build.
 * @param {Article} article - The KB Article to insert into the email body.
 * @param {Date} currentDate - Today's date.
 */
function buildEmailBody(buildType, article, currentDate) {
    var emailBody = "";
    switch(buildType) {
        case BuildEmailBodyTypes.New:
           emailBody = newEmailBody(currentDate);
           break;
        case BuildEmailBodyTypes.Insert:
           emailBody = insertArticleIntoEmailBody(article);
           break;
        case BuildEmailBodyTypes.Conclude:
           emailBody = concludeEmailBody();
           break;
    }
    return emailBody;
}

/**
 * Composes and returns the intro/header of an email body.
 * 
 * @param {Date} currentDate - Today's date.
 */
function newEmailBody(currentDate) {
    var body = "Hello. This email is being sent to inform you that the following TTC KB Articles ";
    body += "will be expiring one week from today on ";
    body += getDateInStringFormat(getNextWeekDate(currentDate)) + ":\r\n\n";
    return body;
}

/**
 * Composes and returns the main content of an email body.
 * 
 * @param {Article} article - The KB Article to insert into the email body.
 */
function insertArticleIntoEmailBody(article) {
    var body = "   - ";
    body += "Article #" + article.id + " which is titled '" + article.title + "' and is owned by ";
    body += article.owner + ".";
    body += "\r\n\n"
    return body;
}

/**
 * Composes and returns the conclusion of an email body.
 */
function concludeEmailBody() {
    var body = "\n";

    body += "STEPS YOU NEED TO TAKE:\r\n";
    body += "   - Open the Tech Team Binder and review the following document:\r\n";
    body += "        - KB Development - Reviewing and Republishing an Expired Article\r\n";
    body += "   - Update the affected articles in the spreadsheet\r\n";
    body += "        - Update any values that may have changed for an article\r\n";
    body += "        - This would include updating the new expiration date\r\n\n";
    
    body += "Please read the spreadsheet's sheet titled 'Instructions & Notes' before any changes are made.\r\n";
    body += "     - The script could malfunction if the spreadsheet is edited improperly.\r\n\n";

    body += "This was an automated email sent from the 'KB Article Expiration Notification' script ";
    body += "which is attached to the 'TTC KB Articles' spreadsheet located in the Google Drive of the ";
    body += "ttctt4@gmail.com email account.\r\n\n";
    body += "If you have any questions or concerns, please email Noah Moss at mossnoah123@gmail.com";
    return body;
}