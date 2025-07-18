const SPREADSHEET_NAME = "Appointment Scheduler Data";
const AVAILABILITY_SHEET_NAME = "Availability";
const ADMIN_EMAIL = "simjinyi@gmail.com";
const FORM_ID_PROPERTY_KEY = "APPOINTMENT_FORM_ID";
const APPOINTMENTS_SHEET_NAME = "Appointments";

const getAvailableSlots = (availabilities, appointments) => {
  const appointmentSet = new Set(appointments);
  return availabilities.filter(
    (availability) => !appointmentSet.has(availability)
  );
};

const setup = () => {
  // get the reference to the form
  const scriptProperties = PropertiesService.getScriptProperties();
  const formId = scriptProperties.getProperty(FORM_ID_PROPERTY_KEY);

  // get the current spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const form = new FormWrapper(spreadsheet, formId);

  setupTriggers(spreadsheet, form);
  updateFormSlots(spreadsheet, form);

  scriptProperties.setProperty(FORM_ID_PROPERTY_KEY, form.getForm().getId());
};

const setupTriggers = (spreadsheet, form) => {
  // delete existing triggers set up by this script
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    // let shouldDelete = false;

    // const handler = trigger.getHandlerFunction();
    // switch (trigger.getEventType()) {
    //   case ScriptApp.EventType.ON_FORM_SUBMIT:
    //     shouldDelete = handler === "onFormSubmit";
    //     break;
    //   case ScriptApp.EventType.ON_EDIT:
    //     shouldDelete = handler === "updateForm";
    //     break;
    // }

    // if (!shouldDelete) return;
    ScriptApp.deleteTrigger(trigger);
  });

  // add the triggers
  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form.getForm())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("updateForm")
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  return true;
};

const updateFormSlots = (spreadsheet, form) => {
  const availabilities = fetchAvailabilities(spreadsheet);
  const appointments = fetchAppointments(spreadsheet);

  const slots = getAvailableSlots(availabilities, appointments);
  form.updateSlots(slots);
};

function updateForm() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const formId = scriptProperties.getProperty(FORM_ID_PROPERTY_KEY);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const form = new FormWrapper(spreadsheet, formId);
  updateFormSlots(spreadsheet, form);
}

function onFormSubmit(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const formId = scriptProperties.getProperty(FORM_ID_PROPERTY_KEY);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const form = new FormWrapper(spreadsheet, formId);

  const availabilitySheet = spreadsheet.getSheetByName(AVAILABILITY_SHEET_NAME);
  if (!availabilitySheet) {
    MailApp.sendEmail(
      ADMIN_EMAIL,
      "Error",
      `availability sheet not found in ${SPREADSHEET_NAME}`
    );
    return;
  }

  const appointmentSheet = spreadsheet.getSheetByName(APPOINTMENTS_SHEET_NAME);
  if (!appointmentSheet) {
    MailApp.sendEmail(
      ADMIN_EMAIL,
      "Error",
      `appointment sheet not found in ${SPREADSHEET_NAME}`
    );
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    MailApp.sendEmail(ADMIN_EMAIL, "Error", "could not obtain lock for script");
    return;
  }

  try {
    const response = e.response;
    const itemResponses = response.getItemResponses();

    let submittedSlot = "";

    itemResponses.forEach((itemResponse) => {
      const answer = itemResponse.getResponse();
      switch (itemResponse.getItem().getTitle()) {
        case SLOT_QUESTION_TITLE:
          submittedSlot = answer;
          break;
      }
    });

    const availabilities = fetchAvailabilities(spreadsheet);
    const availabilitiesSet = new Set(availabilities);
    if (!availabilitiesSet.has(submittedSlot)) {
      MailApp.sendEmail(
        ADMIN_EMAIL,
        "Appointment Booking Failed (Race Condition)",
        `A booking attempt failed due to race condition:\n\nDate: ${bookingDate}\nTime: ${bookingTime}\nAttempted By: ${submittedName} (${submittedEmail})\n\nThis slot was likely booked by another user simultaneously.`
      );

      return;
    }

    appointmentSheet.appendRow([new Date(), submittedSlot]);
    updateFormSlots(spreadsheet, form);

    MailApp.sendEmail(
      ADMIN_EMAIL,
      "Appointment Booking Successful",
      `Booking attempt successful: ${submittedSlot}.`
    );
  } catch (e) {
    MailApp.sendEmail(
      ADMIN_EMAIL,
      "Error",
      `an error occurred during form submission: ${e.toString()}`
    );
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

const isSlotValid = (slot) => {
  const formattedSlot = slot.replace(" - ", " ");

  const dateTime = new Date(formattedSlot);
  if (isNaN(dateTime.getTime())) return true;

  return dateTime.getTime() >= new Date().getTime();
};

const fetchAvailabilities = (spreadsheet) => {
  const sheet = spreadsheet.getSheetByName(AVAILABILITY_SHEET_NAME);
  if (!sheet) throw new Error(ERROR_DEPENDENCY);

  const data = sheet.getDataRange().getValues();
  const filteredData = data
    .filter((_, index) => index > 0)
    .map((row) => row[0])
    .filter(isSlotValid);
  return [...new Set(filteredData)];
};

const fetchAppointments = (spreadsheet) => {
  const sheet = spreadsheet.getSheetByName(APPOINTMENTS_SHEET_NAME);
  if (!sheet) throw ERR_OPEN_SPREADSHEET;

  const data = sheet.getDataRange().getValues();
  return data.filter((_, index) => index > 0).map((row) => row[1]);
};
