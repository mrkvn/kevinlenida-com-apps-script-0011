// CONSTANTS
const SHEET_NAME = 'Sheet1'; // Change if your sheet name is different
const PROCESSED_LABEL_NAME = 'Processed';

function processPldt() {
  const threads = GmailApp.search('from:pldthome@pldt.com.ph subject:"PLDT Electronic Statement dated" has:attachment');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
  const processedLabel = getOrCreateLabel(PROCESSED_LABEL_NAME);

  threads.forEach(thread => {
    if (!threadHasLabel(thread, processedLabel)) {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const body = message.getPlainBody();
        const telephoneNumberMatch = body.match(/(?<=.*Telephone Number.*\s+\n.*\s+\n).*/)
        const statementDateMatch = body.match(/(?<=.*Statement Date.*\s+\n.*\s+\n).*/)
        const accountNameMatch = body.match(/(?<=.*Account Name.*\s+\n.*\s+\n).*/)
        const balanceFromLastBillMatch = body.match(/(?<=.*Balance from Last Bill.*\s+\n.*\s+\n).*/)
        const currentChargesMatch = body.match(/(?<=.*Current Charges.*\s+\n.*\s+\n).*/)
        const dueDateMatch = body.match(/(?<=.*Due Date.*\s+\n.*\s+\n).*/)
        const totalAmountDueMatch = body.match(/(?<=.*Total Amount Due.*\s+\n.*\s+\n).*/)

        if (telephoneNumberMatch && statementDateMatch && accountNameMatch && balanceFromLastBillMatch && currentChargesMatch && dueDateMatch && totalAmountDueMatch) {
          const telephoneNumber = "'" + telephoneNumberMatch[0].replaceAll('*', '')
          const statementDate = formatDate(statementDateMatch[0].replaceAll('*', ''))
          const accountName = accountNameMatch[0].replaceAll('*', '')
          const balanceFromLastBill = balanceFromLastBillMatch[0].replaceAll('*', '')
          const currentCharges = currentChargesMatch[0].replaceAll('*', '')
          const dueDate = formatDate(dueDateMatch[0].replaceAll('*', ''))
          const totalAmountDue = totalAmountDueMatch[0].replaceAll('*', '')

          // Append the extracted data to the sheet
          sheet.appendRow([telephoneNumber, statementDate, accountName, balanceFromLastBill, currentCharges, dueDate, totalAmountDue, new Date()]);

          // Mark the thread as processed by adding the label
          thread.addLabel(processedLabel);
        }
      });
    }
  });
}

function formatDate(dateString) {
  const date = new Date(dateString);
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  return formattedDate
}

function getOrCreateLabel(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  return label ? label : GmailApp.createLabel(labelName);
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}

// To run the script automatically, create a time-driven trigger
function createTrigger() {
  // Delete any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('processPldt')
    .timeBased()
    .after(1000) // immediately run after milliseconds. not exact, will vary
    .create();

  ScriptApp.newTrigger('processPldt')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}
