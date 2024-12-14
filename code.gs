// Declare main parameters
const WHATSAPPCOLUMN = "whatsapp"; // Label from the column of the destination phone number
const STATUSCOLUMN = "status"; // Label for the column where the script will mark the message as processed
const RESCHEDULESTATUSCOLUMN = "Reschedule-status"; // Label for the column where reschedule status is mentioned
const RESCHEDULEDATECOLUMN = "Rescheduled-date"; // Label for rescheduled date
const RESCHEDULETIMECOLUMN = "Rescheduled-time"; // Label for rescheduled time

// WebHook URLs from the 2Chat flow triggers
const WEBHOOK2CHATURL =
  "https://api.p.2chat.io/open/flows/FLW3a72e91f-1b83-4d8d-ba6b-33f1ab557c3c";
const RESCHEDULE_WEBHOOK2CHATURL =
  "https://api.p.2chat.io/open/flows/FLW2be06763-1b73-41c7-b13d-66f4b2a75bd9";

// Function that reads the Google Sheet and sends the messages for new or rescheduled rows
function send2ChatWebhook_Sheet() {
  const sheetName = "contacts";
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusColumnIndex = headers.indexOf(STATUSCOLUMN) + 1;
  const whatsappColumnIndex = headers.indexOf(WHATSAPPCOLUMN) + 1;
  const rescheduleStatusColumnIndex =
    headers.indexOf(RESCHEDULESTATUSCOLUMN) + 1;

  if (
    statusColumnIndex > 0 &&
    whatsappColumnIndex > 0 &&
    rescheduleStatusColumnIndex > 0
  ) {
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[statusColumnIndex - 1];
      const rescheduleStatus = row[rescheduleStatusColumnIndex - 1];

      if (status !== "Processed") {
        send2ChatWebhookRequest(
          sheetName,
          i + 1,
          WHATSAPPCOLUMN,
          STATUSCOLUMN,
          WEBHOOK2CHATURL
        );
      } else if (rescheduleStatus === "YES") {
        send2ChatWebhookRequest(
          sheetName,
          i + 1,
          WHATSAPPCOLUMN,
          STATUSCOLUMN,
          RESCHEDULE_WEBHOOK2CHATURL
        );
        // Clear reschedule status after processing
        sheet.getRange(i + 1, rescheduleStatusColumnIndex).setValue("");
      }
    }
  } else {
    console.error("Required columns not found in the sheet.");
  }
}

// Function to send WhatsApp Message using 2Chat flows trigger
function send2ChatWebhookRequest(
  vSheet,
  vRow,
  vWhatsAppColumn,
  vStatusColumn,
  webhookUrl
) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(vSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet
    .getRange(vRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const JSONrequest = {
    to_number: "",
    variables: {},
  };

  const statusColumnIndex = headers.indexOf(vStatusColumn);

  headers.forEach((header, index) => {
    if (header === vWhatsAppColumn) {
      let phoneNumber = rowData[index].toString().replace(/\D/g, "");
      if (!phoneNumber.startsWith("91")) {
        phoneNumber = "91" + phoneNumber; // Add country code if missing
      }
      JSONrequest.to_number = phoneNumber;
    } else if (header !== vStatusColumn) {
      JSONrequest.variables[header] = rowData[index];
    }
  });

  if (!isValidPhoneNumber(JSONrequest.to_number)) {
    const errorMessage = `Invalid phone number: ${JSONrequest.to_number}`;
    console.error(errorMessage);
    sheet
      .getRange(vRow, statusColumnIndex + 1)
      .setValue("Error: " + errorMessage);
    return;
  }

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(JSONrequest),
    muteHttpExceptions: true, // This allows us to handle the error response
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      sheet.getRange(vRow, statusColumnIndex + 1).setValue("Processed");
    } else {
      const errorMessage = `Error ${responseCode}: ${responseBody}`;
      console.error(errorMessage);
      sheet
        .getRange(vRow, statusColumnIndex + 1)
        .setValue("Error: " + errorMessage);
    }
  } catch (e) {
    const errorMessage = `Error sending webhook request: ${e.message}`;
    console.error(errorMessage);
    sheet
      .getRange(vRow, statusColumnIndex + 1)
      .setValue("Error: " + errorMessage);
  }
}

// Function to validate phone number format for India
function isValidPhoneNumber(phoneNumber) {
  const phoneRegex = /^91\d{10}$/;
  return phoneRegex.test(phoneNumber);
}
