function setup() {
  console.log(CONFIG.LOG_MESSAGES.SETUP_START);
  const userProperties = PropertiesService.getUserProperties();
  let form = initializeForm(userProperties);
  let spreadsheet = initializeSpreadsheet(userProperties, form);

  if (!form || !spreadsheet) {
    console.error(CONFIG.LOG_MESSAGES.SETUP_FAIL);
    return;
  }

  ensureSheetNaming(spreadsheet);
  ensureAvailabilitySheet(spreadsheet);
  setupSpreadsheetTriggers(spreadsheet.getId(), form.getId());

  try {
    updateFormSlots();
    console.log(CONFIG.LOG_MESSAGES.SETUP_UPDATE_SLOTS_SUCCESS);
  } catch (e) {
    console.error(
      `${CONFIG.LOG_MESSAGES.SETUP_UPDATE_SLOTS_ERROR} ${e.message}`
    );
  }

  console.log(CONFIG.LOG_MESSAGES.SETUP_SUCCESS);
  console.log(
    `${CONFIG.LOG_MESSAGES.SETUP_FORM_URL} ${form.getPublishedUrl()}`
  );
  console.log(
    `${CONFIG.LOG_MESSAGES.SETUP_SPREADSHEET_URL} ${spreadsheet.getUrl()}`
  );
}

function onOpenFormHandler() {
  console.log(CONFIG.LOG_MESSAGES.ON_OPEN_TRIGGERED);
  try {
    const userProperties = PropertiesService.getUserProperties();
    const formId = userProperties.getProperty(
      CONFIG.USER_PROPERTIES_KEYS.FORM_ID
    );
    const spreadsheetId = userProperties.getProperty(
      CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID
    );

    if (!formId || !spreadsheetId) {
      console.warn(CONFIG.LOG_MESSAGES.ON_OPEN_WARN_NO_IDS);
      return;
    }
    updateFormSlots();
    console.log(CONFIG.LOG_MESSAGES.ON_OPEN_SUCCESS);
  } catch (error) {
    console.error(`${CONFIG.LOG_MESSAGES.ON_OPEN_ERROR} ${error.message}`);
  }
}

function onEditAppointmentSheet(e) {
  console.log(CONFIG.LOG_MESSAGES.ON_EDIT_TRIGGERED);
  const editedSheetName = e.range.getSheet().getName();

  if (editedSheetName === CONFIG.SHEET_NAMES.AVAILABILITY) {
    console.log(
      CONFIG.LOG_MESSAGES.ON_EDIT_DETECTED_ON_AVAILABILITY.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.AVAILABILITY
      )
    );
    console.log(
      `${CONFIG.LOG_MESSAGES.ON_EDIT_RANGE} ${e.range.getA1Notation()}`
    );
    console.log(`${CONFIG.LOG_MESSAGES.ON_EDIT_NEW_VALUE} ${e.value}`);
    console.log(`${CONFIG.LOG_MESSAGES.ON_EDIT_OLD_VALUE} ${e.oldValue}`);
    console.log(
      `${CONFIG.LOG_MESSAGES.ON_EDIT_USER} ${
        e.user ? e.user.getEmail() : "N/A"
      }`
    );
    console.log(`${CONFIG.LOG_MESSAGES.ON_EDIT_TIMESTAMP} ${new Date()}`);
    updateFormSlots();
    console.log(CONFIG.LOG_MESSAGES.ON_EDIT_UPDATE_SLOTS);
  } else {
    console.log(
      CONFIG.LOG_MESSAGES.ON_EDIT_OTHER_SHEET.replace(
        "{sheetName}",
        editedSheetName
      ).replace("{availabilitySheetName}", CONFIG.SHEET_NAMES.AVAILABILITY)
    );
  }
}

function onFormSubmitAppointmentSheet(e) {
  console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_TRIGGERED);
  const lock = LockService.getScriptLock();
  let form = null;
  let userEmail = null;
  let submissionDetails = {};

  try {
    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_GET_LOCK);
    lock.waitLock(30000);
    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_LOCK_OBTAINED);

    const formResponse = e.response;
    const timestamp = formResponse.getTimestamp();
    const namedValues = e.namedValues || {};

    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_RAW_NAMED_VALUES);
    console.log(JSON.stringify(namedValues, null, 2));

    form = FormApp.openById(e.source.getId());
    form.setConfirmationMessage(CONFIG.CONFIRMATION_MESSAGES.SUCCESS);

    const formItems = form.getItems();
    let submittedSlot = null;
    const headers = ["Timestamp"];
    const rowData = [timestamp];

    formItems.forEach((item) => {
      const itemTitle = item.getTitle();
      headers.push(itemTitle);

      const itemResponse = e.response
        .getItemResponses()
        .find((ir) => ir.getItem().getTitle() === itemTitle);
      const answer = itemResponse ? itemResponse.getResponse() : "";

      if (Array.isArray(answer)) {
        rowData.push(answer.join(", "));
        submissionDetails[itemTitle] = answer.join(", ");
      } else {
        rowData.push(answer);
        submissionDetails[itemTitle] = answer;
      }

      if (itemTitle === CONFIG.FORM_TITLES.TIME_SLOT) {
        submittedSlot = answer;
      }
      if (itemTitle === CONFIG.FORM_TITLES.EMAIL) {
        userEmail = answer;
      }
    });

    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_DETAILS);
    console.log(`${CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_TIMESTAMP} ${timestamp}`);
    console.log(
      `${
        CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_RESPONSE_ID
      } ${formResponse.getId()}`
    );
    console.log(
      `${CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_SUBMITTED_BY} ${userEmail || "N/A"}`
    );
    console.log(
      `${CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_SUBMITTED_SLOT} ${submittedSlot}`
    );
    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_ANSWERS);
    headers.forEach((header, index) => {
      console.log(`${header}: ${rowData[index]}`);
    });

    const userProperties = PropertiesService.getUserProperties();
    const spreadsheetId = userProperties.getProperty(
      CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID
    );

    if (!spreadsheetId) {
      console.error(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_ERROR_NO_SS_ID);
      sendConfirmationEmail(
        CONFIG.ADMIN_EMAIL,
        CONFIG.EMAIL_MESSAGES.ADMIN_SUBJECT_NO_SS_ID,
        CONFIG.EMAIL_MESSAGES.ADMIN_BODY_NO_SS_ID.replace(
          "{submissionDetails}",
          JSON.stringify(submissionDetails)
        ),
        null
      );
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const appointmentSheet = spreadsheet.getSheetByName(
      CONFIG.SHEET_NAMES.APPOINTMENTS
    );

    if (!appointmentSheet) {
      console.error(
        CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_ERROR_SHEET_NOT_FOUND.replace(
          "{sheetName}",
          CONFIG.SHEET_NAMES.APPOINTMENTS
        ).replace("{spreadsheetId}", spreadsheetId)
      );
      sendConfirmationEmail(
        CONFIG.ADMIN_EMAIL,
        CONFIG.EMAIL_MESSAGES.ADMIN_SUBJECT_SHEET_NOT_FOUND,
        CONFIG.EMAIL_MESSAGES.ADMIN_BODY_SHEET_NOT_FOUND.replace(
          "{sheetName}",
          CONFIG.SHEET_NAMES.APPOINTMENTS
        )
          .replace("{spreadsheetId}", spreadsheetId)
          .replace("{submissionDetails}", JSON.stringify(submissionDetails)),
        userEmail,
        CONFIG.EMAIL_MESSAGES.USER_BODY_ERROR,
        CONFIG.EMAIL_MESSAGES.USER_SUBJECT_ERROR
      );
      return;
    }

    const existingAppointments = fetchAppointments(spreadsheet);
    const appointmentSet = new Set(existingAppointments);

    if (appointmentSet.has(submittedSlot)) {
      console.warn(
        CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_DUPLICATE_DETECTED.replace(
          "{submittedSlot}",
          submittedSlot
        )
      );
      form.setConfirmationMessage(CONFIG.CONFIRMATION_MESSAGES.DUPLICATE);
      sendConfirmationEmail(
        CONFIG.ADMIN_EMAIL,
        CONFIG.EMAIL_MESSAGES.ADMIN_SUBJECT_DUPLICATE,
        CONFIG.EMAIL_MESSAGES.ADMIN_BODY_DUPLICATE.replace(
          "{submittedSlot}",
          submittedSlot
        )
          .replace("{userEmail}", userEmail || "N/A")
          .replace("{submissionDetails}", JSON.stringify(submissionDetails)),
        userEmail,
        CONFIG.EMAIL_MESSAGES.USER_SUBJECT_DUPLICATE,
        CONFIG.CONFIRMATION_MESSAGES.DUPLICATE
      );
      return;
    }

    if (appointmentSheet.getLastRow() === 0) {
      appointmentSheet.appendRow(headers);
      console.log(
        CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_ADDED_HEADERS.replace(
          "{sheetName}",
          CONFIG.SHEET_NAMES.APPOINTMENTS
        )
      );
    }

    appointmentSheet.appendRow(rowData);
    console.log(
      CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_APPENDED_ROW.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.APPOINTMENTS
      )
    );

    console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_UPDATING_SLOTS);
    updateFormSlots();

    sendConfirmationEmail(
      CONFIG.ADMIN_EMAIL,
      CONFIG.EMAIL_MESSAGES.ADMIN_SUBJECT_NEW_APPOINTMENT.replace(
        "{submittedSlot}",
        submittedSlot
      ),
      CONFIG.EMAIL_MESSAGES.ADMIN_BODY_NEW_APPOINTMENT.replace(
        "{submissionDetails}",
        formatSubmissionDetails(submissionDetails)
      ),
      userEmail,
      CONFIG.EMAIL_MESSAGES.USER_SUBJECT_SUCCESS.replace(
        "{submittedSlot}",
        submittedSlot
      ),
      `${CONFIG.EMAIL_MESSAGES.USER_BODY_SUCCESS_PART1.replace(
        "{userName}",
        submissionDetails[CONFIG.FORM_TITLES.NAME] || "Valued Patient"
      ).replace("{submittedSlot}", submittedSlot)}${formatSubmissionDetails(
        submissionDetails
      )}${CONFIG.EMAIL_MESSAGES.USER_BODY_SUCCESS_PART2}`
    );
  } catch (error) {
    console.error(
      `${CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_PROCESSING_ERROR} ${error.message}`
    );
    if (form) {
      form.setConfirmationMessage(CONFIG.CONFIRMATION_MESSAGES.ERROR);
    }
    sendConfirmationEmail(
      CONFIG.ADMIN_EMAIL,
      CONFIG.EMAIL_MESSAGES.ADMIN_SUBJECT_SUBMISSION_ERROR,
      CONFIG.EMAIL_MESSAGES.ADMIN_BODY_SUBMISSION_ERROR.replace(
        "{errorMessage}",
        error.message
      ).replace("{submissionDetails}", JSON.stringify(submissionDetails)),
      userEmail,
      CONFIG.EMAIL_MESSAGES.USER_SUBJECT_ERROR,
      CONFIG.CONFIRMATION_MESSAGES.ERROR
    );
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
      console.log(CONFIG.LOG_MESSAGES.ON_FORM_SUBMIT_LOCK_RELEASED);
    }
  }
}

function updateFormSlots() {
  console.log(CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_START);
  const userProperties = PropertiesService.getUserProperties();
  const formId = userProperties.getProperty(
    CONFIG.USER_PROPERTIES_KEYS.FORM_ID
  );
  const spreadsheetId = userProperties.getProperty(
    CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID
  );

  if (!formId) {
    console.warn(CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_WARN_NO_FORM_ID);
    return;
  }
  let form = FormApp.openById(formId);

  if (!spreadsheetId) {
    console.warn(CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_WARN_NO_SS_ID);
    return;
  }
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  try {
    const availabilities = fetchAvailabilities(spreadsheet);
    const appointments = fetchAppointments(spreadsheet);

    console.log(
      `${
        CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_AVAILABILITIES_FETCHED
      } ${availabilities.join(", ")}`
    );
    console.log(
      `${
        CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_APPOINTMENTS_FETCHED
      } ${appointments.join(", ")}`
    );

    const availableSlots = getAvailableSlots(availabilities, appointments);
    const slotQuestion = findFormItemByTitle(
      form,
      CONFIG.FORM_TITLES.TIME_SLOT
    );

    if (!slotQuestion) {
      console.error(CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_NO_SLOT_QUESTION);
      return;
    }

    const sortedSlots = availableSlots.sort(slotSorterFn);
    console.log(
      `${CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_SORTED} ${sortedSlots.join(
        ", "
      )}`
    );

    slotQuestion
      .asListItem()
      .setChoiceValues(
        sortedSlots.length > 0 ? sortedSlots : [CONFIG.DEFAULT_SLOT_MESSAGE]
      );
    console.log(CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_SUCCESS);
  } catch (e) {
    console.error(
      `${CONFIG.LOG_MESSAGES.UPDATE_FORM_SLOTS_ERROR} ${e.message}`
    );
  }
}

function initializeForm(userProperties) {
  console.log(CONFIG.LOG_MESSAGES.INIT_FORM_START);
  let formId = userProperties.getProperty(CONFIG.USER_PROPERTIES_KEYS.FORM_ID);
  let form = null;

  if (formId) {
    try {
      form = FormApp.openById(formId);
      console.log(
        `${CONFIG.LOG_MESSAGES.INIT_FORM_EXISTING_FOUND.replace(
          "{formId}",
          formId
        )}`
      );
    } catch (e) {
      console.warn(
        `${CONFIG.LOG_MESSAGES.INIT_FORM_ERROR_OPENING.replace(
          "{formId}",
          formId
        )} ${e.message}`
      );
      userProperties.deleteProperty(CONFIG.USER_PROPERTIES_KEYS.FORM_ID);
      formId = null;
    }
  }

  if (!form) {
    console.log(CONFIG.LOG_MESSAGES.INIT_FORM_CREATING_NEW);
    form = FormApp.create(CONFIG.FORM_TITLES.MAIN);
    form.setTitle(CONFIG.FORM_TITLES.MAIN);
    form.setDescription(CONFIG.FORM_DESCRIPTIONS.MAIN);

    form.addTextItem().setTitle(CONFIG.FORM_TITLES.NAME).setRequired(true);

    const emailValidation = FormApp.createTextValidation()
      .requireTextMatchesPattern(CONFIG.REGEX_PATTERNS.EMAIL_VALIDATION)
      .setHelpText(CONFIG.EMAIL_MESSAGES.USER_BODY_ERROR) // Reusing error message for help text
      .build();

    form
      .addTextItem()
      .setTitle(CONFIG.FORM_TITLES.EMAIL)
      .setRequired(false)
      .setValidation(emailValidation);

    form
      .addTextItem()
      .setTitle(CONFIG.FORM_TITLES.PHONE_NUMBER)
      .setRequired(true);

    form
      .addListItem()
      .setTitle(CONFIG.FORM_TITLES.TIME_SLOT)
      .setRequired(true)
      .setChoiceValues([CONFIG.DEFAULT_SLOT_MESSAGE]);

    formId = form.getId();
    userProperties.setProperty(CONFIG.USER_PROPERTIES_KEYS.FORM_ID, formId);
    console.log(
      `${CONFIG.LOG_MESSAGES.INIT_FORM_CREATED_NEW.replace("{formId}", formId)}`
    );
    console.log(
      `${CONFIG.LOG_MESSAGES.INIT_FORM_URL} ${form.getPublishedUrl()}`
    );
  }
  return form;
}

function initializeSpreadsheet(userProperties, form) {
  console.log(CONFIG.LOG_MESSAGES.INIT_SS_START);
  let spreadsheetId = userProperties.getProperty(
    CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID
  );
  let spreadsheet = null;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log(
        `${CONFIG.LOG_MESSAGES.INIT_SS_EXISTING_LINKED.replace(
          "{spreadsheetId}",
          spreadsheetId
        )}`
      );
    } catch (e) {
      console.warn(
        `${CONFIG.LOG_MESSAGES.INIT_SS_ERROR_OPENING.replace(
          "{spreadsheetId}",
          spreadsheetId
        )} ${e.message}`
      );
      userProperties.deleteProperty(CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID);
      spreadsheetId = null;
    }
  }

  if (!spreadsheet) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSpreadsheet) {
      spreadsheet = activeSpreadsheet;
      spreadsheetId = spreadsheet.getId();
      userProperties.setProperty(
        CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID,
        spreadsheetId
      );
      console.log(
        `${CONFIG.LOG_MESSAGES.INIT_SS_USING_ACTIVE.replace(
          "{spreadsheetId}",
          spreadsheetId
        )}`
      );
    } else {
      console.log(CONFIG.LOG_MESSAGES.INIT_SS_NO_ACTIVE);
      spreadsheet = SpreadsheetApp.create(`Responses for '${form.getTitle()}'`);
      spreadsheetId = spreadsheet.getId();
      userProperties.setProperty(
        CONFIG.USER_PROPERTIES_KEYS.SPREADSHEET_ID,
        spreadsheetId
      );
      console.log(
        `${CONFIG.LOG_MESSAGES.INIT_SS_CREATED_NEW.replace(
          "{spreadsheetId}",
          spreadsheetId
        )}`
      );
    }
  }
  return spreadsheet;
}

function ensureSheetNaming(spreadsheet) {
  console.log(CONFIG.LOG_MESSAGES.ENSURE_SHEET_NAMING_START);
  const formResponsesSheet = spreadsheet.getSheets()[0];
  if (!formResponsesSheet) {
    console.error(CONFIG.LOG_MESSAGES.ENSURE_SHEET_NAMING_NO_SHEET);
    return;
  }
  const originalName = formResponsesSheet.getName();
  if (originalName !== CONFIG.SHEET_NAMES.APPOINTMENTS) {
    formResponsesSheet.setName(CONFIG.SHEET_NAMES.APPOINTMENTS);
    console.log(
      CONFIG.LOG_MESSAGES.ENSURE_SHEET_NAMING_RENAMED.replace(
        "{originalName}",
        originalName
      ).replace("{newName}", CONFIG.SHEET_NAMES.APPOINTMENTS)
    );
  } else {
    console.log(
      CONFIG.LOG_MESSAGES.ENSURE_SHEET_NAMING_ALREADY_NAMED.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.APPOINTMENTS
      )
    );
  }
}

function ensureAvailabilitySheet(spreadsheet) {
  console.log(CONFIG.LOG_MESSAGES.ENSURE_AVAILABILITY_SHEET_START);
  let availabilitySheet = spreadsheet.getSheetByName(
    CONFIG.SHEET_NAMES.AVAILABILITY
  );
  if (!availabilitySheet) {
    availabilitySheet = spreadsheet.insertSheet(
      CONFIG.SHEET_NAMES.AVAILABILITY
    );
    availabilitySheet.appendRow(["Available Time Slot"]); // This is a fixed header, not user-configurable text.
    console.log(
      CONFIG.LOG_MESSAGES.ENSURE_AVAILABILITY_SHEET_CREATED.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.AVAILABILITY
      )
    );
  } else {
    console.log(
      CONFIG.LOG_MESSAGES.ENSURE_AVAILABILITY_SHEET_EXISTS.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.AVAILABILITY
      )
    );
  }
}

function setupSpreadsheetTriggers(spreadsheetId, formId) {
  console.log(CONFIG.LOG_MESSAGES.SETUP_TRIGGERS_START);
  const triggers = ScriptApp.getProjectTriggers();
  const triggerFunctions = Object.values(CONFIG.TRIGGER_FUNCTIONS);

  for (const trigger of triggers) {
    const handler = trigger.getHandlerFunction();
    if (triggerFunctions.includes(handler)) {
      ScriptApp.deleteTrigger(trigger);
      console.log(
        CONFIG.LOG_MESSAGES.SETUP_TRIGGERS_DELETED_EXISTING.replace(
          "{handler}",
          handler
        )
      );
    }
  }

  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTIONS.ON_EDIT)
    .forSpreadsheet(spreadsheetId)
    .onEdit()
    .create();
  console.log(
    CONFIG.LOG_MESSAGES.SETUP_TRIGGERS_ON_EDIT_CREATED.replace(
      "{spreadsheetId}",
      spreadsheetId
    )
  );

  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTIONS.ON_FORM_SUBMIT)
    .forForm(formId)
    .onFormSubmit()
    .create();
  console.log(
    CONFIG.LOG_MESSAGES.SETUP_TRIGGERS_ON_FORM_SUBMIT_CREATED.replace(
      "{formId}",
      formId
    )
  );

  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTIONS.ON_FORM_OPEN)
    .forForm(formId)
    .onOpen()
    .create();
  console.log(
    CONFIG.LOG_MESSAGES.SETUP_TRIGGERS_ON_OPEN_CREATED.replace(
      "{formId}",
      formId
    )
  );
}

function isSlotValid(slot) {
  console.log(
    CONFIG.LOG_MESSAGES.IS_SLOT_VALID_CHECKING.replace("{slot}", slot)
  );
  if (typeof slot !== "string" || slot.trim() === "") {
    console.log(
      CONFIG.LOG_MESSAGES.IS_SLOT_VALID_NOT_STRING_OR_EMPTY.replace(
        "{slot}",
        slot
      )
    );
    return false;
  }
  const formattedSlot = slot.replace(" - ", " ");
  const dateTime = new Date(formattedSlot);
  if (isNaN(dateTime.getTime())) {
    console.log(
      CONFIG.LOG_MESSAGES.IS_SLOT_VALID_CANNOT_PARSE_DATE.replace(
        "{slot}",
        slot
      )
    );
    return true;
  }
  const isValid = dateTime.getTime() >= new Date().getTime();
  console.log(
    CONFIG.LOG_MESSAGES.IS_SLOT_VALID_RESULT.replace("{slot}", slot).replace(
      "{validity}",
      isValid ? "valid" : "invalid"
    )
  );
  return isValid;
}

function fetchAvailabilities(spreadsheet) {
  console.log(CONFIG.LOG_MESSAGES.FETCH_AVAILABILITIES_START);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.AVAILABILITY);
  if (!sheet) {
    console.error(
      CONFIG.LOG_MESSAGES.FETCH_AVAILABILITIES_SHEET_NOT_FOUND.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.AVAILABILITY
      )
    );
    throw new Error(
      `Dependent sheet "${CONFIG.SHEET_NAMES.AVAILABILITY}" not found.`
    );
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log(CONFIG.LOG_MESSAGES.FETCH_AVAILABILITIES_EMPTY_OR_HEADERS);
    return [];
  }
  const filteredData = data
    .filter((_, index) => index > 0)
    .map((row) => row[0])
    .filter((slot) => isSlotValid(slot));

  const uniqueSlots = [...new Set(filteredData)];
  console.log(
    CONFIG.LOG_MESSAGES.FETCH_AVAILABILITIES_FOUND.replace(
      "{count}",
      uniqueSlots.length
    )
  );
  return uniqueSlots;
}

const fetchAppointments = (spreadsheet) => {
  console.log(CONFIG.LOG_MESSAGES.FETCH_APPOINTMENTS_START);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.APPOINTMENTS);
  if (!sheet) {
    console.error(
      CONFIG.LOG_MESSAGES.FETCH_APPOINTMENTS_SHEET_NOT_FOUND.replace(
        "{sheetName}",
        CONFIG.SHEET_NAMES.APPOINTMENTS
      )
    );
    throw new Error(
      `Could not open the linked spreadsheet or sheet "${CONFIG.SHEET_NAMES.APPOINTMENTS}".`
    );
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log(CONFIG.LOG_MESSAGES.FETCH_APPOINTMENTS_EMPTY_OR_HEADERS);
    return [];
  }

  const headers = data[0];
  const timeSlotColumnIndex = headers.indexOf(CONFIG.FORM_TITLES.TIME_SLOT);
  if (timeSlotColumnIndex === -1) {
    console.warn(CONFIG.LOG_MESSAGES.FETCH_APPOINTMENTS_NO_TIME_SLOT_COLUMN);
    return [];
  }

  const appointments = data
    .filter((_, index) => index > 0)
    .map((row) => row[timeSlotColumnIndex])
    .filter((slot) => typeof slot === "string" && slot.trim() !== "");
  console.log(
    CONFIG.LOG_MESSAGES.FETCH_APPOINTMENTS_FOUND.replace(
      "{count}",
      appointments.length
    )
  );
  return appointments;
};

function findFormItemByTitle(form, title) {
  console.log(
    CONFIG.LOG_MESSAGES.FIND_FORM_ITEM_START.replace("{title}", title)
  );
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) {
      console.log(
        CONFIG.LOG_MESSAGES.FIND_FORM_ITEM_FOUND.replace("{title}", title)
      );
      return items[i];
    }
  }
  console.warn(
    CONFIG.LOG_MESSAGES.FIND_FORM_ITEM_NOT_FOUND.replace("{title}", title)
  );
  return null;
}

function slotSorterFn(a, b) {
  const dateA = new Date(a.replace(" - ", " "));
  const dateB = new Date(b.replace(" - ", " "));

  const isValidA = !isNaN(dateA.getTime());
  const isValidB = !isNaN(dateB.getTime());

  if (isValidA && isValidB) {
    return dateA.getTime() - dateB.getTime();
  }

  if (isValidA && !isValidB) {
    return -1;
  }

  if (!isValidA && isValidB) {
    return 1;
  }

  if (!isValidA && !isValidB) {
    return a.localeCompare(b);
  }
  return 0;
}

function getAvailableSlots(availabilities, appointments) {
  console.log(CONFIG.LOG_MESSAGES.GET_AVAILABLE_SLOTS_START);
  const appointmentSet = new Set(appointments);
  const available = availabilities.filter(
    (availability) => !appointmentSet.has(availability)
  );
  console.log(
    CONFIG.LOG_MESSAGES.GET_AVAILABLE_SLOTS_FOUND.replace(
      "{count}",
      available.length
    )
  );
  return available;
}

function getAppointmentSheetHeaders() {
  return ["Timestamp", "Name", "Email", "Phone Number", "Time Slot"];
}

function sendConfirmationEmail(
  adminEmail,
  adminSubject,
  adminBody,
  userEmail = null,
  userSubject = null,
  userBody = null
) {
  console.log(CONFIG.LOG_MESSAGES.SEND_EMAIL_ATTEMPT);
  try {
    MailApp.sendEmail(adminEmail, adminSubject, adminBody);
    console.log(
      CONFIG.LOG_MESSAGES.SEND_EMAIL_ADMIN_SUCCESS.replace(
        "{email}",
        adminEmail
      ).replace("{subject}", adminSubject)
    );
  } catch (error) {
    console.error(
      CONFIG.LOG_MESSAGES.SEND_EMAIL_ADMIN_FAIL.replace(
        "{email}",
        adminEmail
      ).replace("{error}", error.message)
    );
  }

  if (userEmail && userSubject && userBody) {
    try {
      MailApp.sendEmail(userEmail, userSubject, userBody);
      console.log(
        CONFIG.LOG_MESSAGES.SEND_EMAIL_USER_SUCCESS.replace(
          "{email}",
          userEmail
        ).replace("{subject}", userSubject)
      );
    } catch (error) {
      console.error(
        CONFIG.LOG_MESSAGES.SEND_EMAIL_USER_FAIL.replace(
          "{email}",
          userEmail
        ).replace("{error}", error.message)
      );
    }
  } else {
    console.log(CONFIG.LOG_MESSAGES.SEND_EMAIL_SKIP_USER);
  }
}

function formatSubmissionDetails(details) {
  let formattedString = "";
  for (const key in details) {
    if (Object.prototype.hasOwnProperty.call(details, key)) {
      formattedString += `- ${key}: ${details[key]}\n`;
    }
  }
  return formattedString;
}
