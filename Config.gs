const CONFIG = {
  USER_PROPERTIES_KEYS: {
    FORM_ID: "formID",
    SPREADSHEET_ID: "spreadsheetID",
  },
  SHEET_NAMES: {
    APPOINTMENTS: "Appointment",
    AVAILABILITY: "Availability",
  },
  DEFAULT_SLOT_MESSAGE: "No slots available. Please check back later.",
  FORM_TITLES: {
    MAIN: "Dental Appointment Form",
    NAME: "Name",
    EMAIL: "Email",
    PHONE_NUMBER: "Phone Number",
    TIME_SLOT: "Time Slot",
  },
  FORM_DESCRIPTIONS: {
    MAIN: "Make an appointment for dental services.",
  },
  REGEX_PATTERNS: {
    EMAIL_VALIDATION: "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$",
  },
  CONFIRMATION_MESSAGES: {
    SUCCESS:
      "Thank you for your submission! Your responses have been recorded.",
    DUPLICATE:
      "It looks like that time slot has just been taken. Please choose another available slot. 😟",
    ERROR: "There was an error processing your request. Please try again. 😔",
  },
  TRIGGER_FUNCTIONS: {
    ON_EDIT: "onEditAppointmentSheet",
    ON_FORM_SUBMIT: "onFormSubmitAppointmentSheet",
    ON_FORM_OPEN: "onOpenFormHandler",
  },
  ADMIN_EMAIL: "your.admin.email@example.com", // 📧 IMPORTANT: Replace with your actual admin email address!

  // --- ALL EXTRACTED TEXT STRINGS ---
  LOG_MESSAGES: {
    SETUP_START: "setup: Starting script setup process.",
    SETUP_FAIL:
      "setup: Failed to initialize form or spreadsheet. Exiting setup.",
    SETUP_SUCCESS: "setup: Script setup complete.",
    SETUP_FORM_URL: "setup: Form URL:",
    SETUP_SPREADSHEET_URL: "setup: Spreadsheet URL:",
    SETUP_UPDATE_SLOTS_SUCCESS:
      "setup: Initial form slots updated successfully.",
    SETUP_UPDATE_SLOTS_ERROR: "setup: Error updating form slots during setup:",

    ON_OPEN_TRIGGERED: "onOpenFormHandler: Triggered.",
    ON_OPEN_WARN_NO_IDS:
      "onOpenFormHandler: Form ID or Spreadsheet ID not found in UserProperties. Cannot update form slots.",
    ON_OPEN_SUCCESS: "onOpenFormHandler: Form slots updated successfully.",
    ON_OPEN_ERROR: "onOpenFormHandler: Error:",

    ON_EDIT_TRIGGERED: "onEditAppointmentSheet: Triggered.",
    ON_EDIT_DETECTED_ON_AVAILABILITY:
      'onEditAppointmentSheet: Edit detected on sheet "{sheetName}":',
    ON_EDIT_RANGE: "onEditAppointmentSheet: Range:",
    ON_EDIT_NEW_VALUE: "onEditAppointmentSheet: New Value:",
    ON_EDIT_OLD_VALUE: "onEditAppointmentSheet: Old Value:",
    ON_EDIT_USER: "onEditAppointmentSheet: User:",
    ON_EDIT_TIMESTAMP: "onEditAppointmentSheet: Timestamp:",
    ON_EDIT_UPDATE_SLOTS:
      "onEditAppointmentSheet: Form slots updated due to availability sheet edit.",
    ON_EDIT_OTHER_SHEET:
      'onEditAppointmentSheet: Edit detected on sheet "{sheetName}", not "{availabilitySheetName}". No specific action taken.',

    ON_FORM_SUBMIT_TRIGGERED: "onFormSubmitAppointmentSheet: Triggered.",
    ON_FORM_SUBMIT_GET_LOCK:
      "onFormSubmitAppointmentSheet: Attempting to get lock...",
    ON_FORM_SUBMIT_LOCK_OBTAINED:
      "onFormSubmitAppointmentSheet: Lock obtained.",
    ON_FORM_SUBMIT_RAW_NAMED_VALUES:
      "onFormSubmitAppointmentSheet: Raw e.namedValues object:",
    ON_FORM_SUBMIT_DETAILS:
      "onFormSubmitAppointmentSheet: Form Submission details:",
    ON_FORM_SUBMIT_TIMESTAMP: "onFormSubmitAppointmentSheet: Timestamp:",
    ON_FORM_SUBMIT_RESPONSE_ID: "onFormSubmitAppointmentSheet: Response ID:",
    ON_FORM_SUBMIT_SUBMITTED_BY:
      "onFormSubmitAppointmentSheet: Submitted by (Email):",
    ON_FORM_SUBMIT_SUBMITTED_SLOT:
      "onFormSubmitAppointmentSheet: Submitted Slot:",
    ON_FORM_SUBMIT_ANSWERS:
      "onFormSubmitAppointmentSheet: Answers (Dynamically Extracted):",
    ON_FORM_SUBMIT_ERROR_NO_SS_ID:
      "onFormSubmitAppointmentSheet: Error: No linked spreadsheet ID found in UserProperties for manual insertion.",
    ON_FORM_SUBMIT_ERROR_SHEET_NOT_FOUND:
      'onFormSubmitAppointmentSheet: Error: Sheet "{sheetName}" not found in spreadsheet ID: {spreadsheetId}',
    ON_FORM_SUBMIT_DUPLICATE_DETECTED:
      'onFormSubmitAppointmentSheet: Duplicate submission detected for slot: "{submittedSlot}". Not processing.',
    ON_FORM_SUBMIT_ADDED_HEADERS:
      'onFormSubmitAppointmentSheet: Added dynamic header row to "{sheetName}" sheet.',
    ON_FORM_SUBMIT_APPENDED_ROW:
      'onFormSubmitAppointmentSheet: Manually appended form response to "{sheetName}" sheet.',
    ON_FORM_SUBMIT_UPDATING_SLOTS:
      "onFormSubmitAppointmentSheet: Updating form slots to reflect booked appointment.",
    ON_FORM_SUBMIT_PROCESSING_ERROR:
      "onFormSubmitAppointmentSheet: Error during form submission processing:",
    ON_FORM_SUBMIT_LOCK_RELEASED:
      "onFormSubmitAppointmentSheet: Lock released.",

    UPDATE_FORM_SLOTS_START: "updateFormSlots: Starting update of form slots.",
    UPDATE_FORM_SLOTS_WARN_NO_FORM_ID:
      "updateFormSlots: Form ID not found in user properties. Cannot update form slots.",
    UPDATE_FORM_SLOTS_WARN_NO_SS_ID:
      "updateFormSlots: Spreadsheet ID not found in user properties. Cannot update form slots.",
    UPDATE_FORM_SLOTS_AVAILABILITIES_FETCHED:
      "updateFormSlots: Availabilities fetched:",
    UPDATE_FORM_SLOTS_APPOINTMENTS_FETCHED:
      "updateFormSlots: Appointments fetched:",
    UPDATE_FORM_SLOTS_NO_SLOT_QUESTION:
      "updateFormSlots: Time Slot question not found in the form. Cannot update slots.",
    UPDATE_FORM_SLOTS_SORTED: "updateFormSlots: Sorted available slots:",
    UPDATE_FORM_SLOTS_SUCCESS:
      "updateFormSlots: Form slots updated successfully.",
    UPDATE_FORM_SLOTS_ERROR: "updateFormSlots: Error during slot update:",

    INIT_FORM_START: "initializeForm: Attempting to initialize form.",
    INIT_FORM_EXISTING_FOUND:
      "initializeForm: Found existing form with ID: {formId}",
    INIT_FORM_ERROR_OPENING:
      "initializeForm: Could not open form with ID {formId}. It might have been deleted. Error:",
    INIT_FORM_CREATING_NEW:
      "initializeForm: No existing form found or could not open. Creating new form.",
    INIT_FORM_CREATED_NEW: "initializeForm: Created new form with ID: {formId}",
    INIT_FORM_URL: "initializeForm: Form URL:",

    INIT_SS_START:
      "initializeSpreadsheet: Attempting to initialize spreadsheet.",
    INIT_SS_EXISTING_LINKED:
      "initializeSpreadsheet: Found existing linked spreadsheet with ID: {spreadsheetId}",
    INIT_SS_ERROR_OPENING:
      "initializeSpreadsheet: Could not open stored spreadsheet with ID {spreadsheetId}. It might have been deleted. Error:",
    INIT_SS_USING_ACTIVE:
      "initializeSpreadsheet: Using active spreadsheet as response destination. ID: {spreadsheetId}",
    INIT_SS_NO_ACTIVE:
      "initializeSpreadsheet: No active spreadsheet. Creating new spreadsheet for responses.",
    INIT_SS_CREATED_NEW:
      "initializeSpreadsheet: Created new spreadsheet for responses. ID: {spreadsheetId}",

    ENSURE_SHEET_NAMING_START:
      "ensureSheetNaming: Ensuring form responses sheet is correctly named.",
    ENSURE_SHEET_NAMING_NO_SHEET:
      "ensureSheetNaming: Could not find the form responses sheet.",
    ENSURE_SHEET_NAMING_RENAMED:
      'ensureSheetNaming: Renamed form responses sheet from "{originalName}" to "{newName}"',
    ENSURE_SHEET_NAMING_ALREADY_NAMED:
      'ensureSheetNaming: Form responses sheet is already named "{sheetName}"',

    ENSURE_AVAILABILITY_SHEET_START:
      "ensureAvailabilitySheet: Ensuring 'Availability' sheet exists.",
    ENSURE_AVAILABILITY_SHEET_CREATED:
      'ensureAvailabilitySheet: Created new sheet "{sheetName}" for availability.',
    ENSURE_AVAILABILITY_SHEET_EXISTS:
      'ensureAvailabilitySheet: Sheet "{sheetName}" already exists.',

    SETUP_TRIGGERS_START:
      "setupSpreadsheetTriggers: Setting up spreadsheet and form triggers.",
    SETUP_TRIGGERS_DELETED_EXISTING:
      "setupSpreadsheetTriggers: Deleted existing trigger for function: {handler}",
    SETUP_TRIGGERS_ON_EDIT_CREATED:
      "setupSpreadsheetTriggers: Created new onEdit trigger for spreadsheet ID {spreadsheetId}.",
    SETUP_TRIGGERS_ON_FORM_SUBMIT_CREATED:
      "setupSpreadsheetTriggers: Created new onFormSubmit trigger for form ID {formId}.",
    SETUP_TRIGGERS_ON_OPEN_CREATED:
      "setupSpreadsheetTriggers: Created new onOpen trigger for form ID {formId}.",

    IS_SLOT_VALID_CHECKING: "isSlotValid: Checking validity for slot: {slot}",
    IS_SLOT_VALID_NOT_STRING_OR_EMPTY:
      'isSlotValid: Slot "{slot}" is not a valid string or is empty.',
    IS_SLOT_VALID_CANNOT_PARSE_DATE:
      'isSlotValid: Slot "{slot}" cannot be parsed as a date. Considering valid for now.',
    IS_SLOT_VALID_RESULT:
      'isSlotValid: Slot "{slot}" is {validity} based on current time.',

    FETCH_AVAILABILITIES_START:
      "fetchAvailabilities: Fetching available time slots.",
    FETCH_AVAILABILITIES_SHEET_NOT_FOUND:
      'fetchAvailabilities: Error: "{sheetName}" sheet not found.',
    FETCH_AVAILABILITIES_EMPTY_OR_HEADERS:
      "fetchAvailabilities: Availability sheet is empty or only contains headers.",
    FETCH_AVAILABILITIES_FOUND:
      "fetchAvailabilities: Found {count} unique and valid availabilities.",

    FETCH_APPOINTMENTS_START:
      "fetchAppointments: Fetching booked appointments.",
    FETCH_APPOINTMENTS_SHEET_NOT_FOUND:
      'fetchAppointments: Error: "{sheetName}" sheet not found.',
    FETCH_APPOINTMENTS_EMPTY_OR_HEADERS:
      "fetchAppointments: Appointment sheet is empty or only contains headers.",
    FETCH_APPOINTMENTS_NO_TIME_SLOT_COLUMN:
      "fetchAppointments: 'Time Slot' column not found in Appointment sheet headers. Check data extraction logic.",
    FETCH_APPOINTMENTS_FOUND: "fetchAppointments: Found {count} appointments.",

    FIND_FORM_ITEM_START:
      'findFormItemByTitle: Searching for form item with title: "{title}".',
    FIND_FORM_ITEM_FOUND: 'findFormItemByTitle: Found item "{title}".',
    FIND_FORM_ITEM_NOT_FOUND:
      'findFormItemByTitle: Item with title "{title}" not found.',

    GET_AVAILABLE_SLOTS_START:
      "getAvailableSlots: Calculating available slots.",
    GET_AVAILABLE_SLOTS_FOUND:
      "getAvailableSlots: Found {count} available slots.",

    SEND_EMAIL_ATTEMPT: "sendConfirmationEmail: Attempting to send emails.",
    SEND_EMAIL_ADMIN_SUCCESS:
      'sendConfirmationEmail: Admin email sent to {email} with subject: "{subject}".',
    SEND_EMAIL_ADMIN_FAIL:
      "sendConfirmationEmail: Failed to send admin email to {email}: {error}.",
    SEND_EMAIL_USER_SUCCESS:
      'sendConfirmationEmail: User email sent to {email} with subject: "{subject}".',
    SEND_EMAIL_USER_FAIL:
      "sendConfirmationEmail: Failed to send user email to {email}: {error}.",
    SEND_EMAIL_SKIP_USER:
      "sendConfirmationEmail: No user email, subject, or body provided. Skipping user email.",
  },
  EMAIL_MESSAGES: {
    ADMIN_SUBJECT_NO_SS_ID:
      "Appointment Form Submission Error - No Spreadsheet ID",
    ADMIN_BODY_NO_SS_ID:
      "A form was submitted, but no linked spreadsheet ID was found. Form details: {submissionDetails}",
    ADMIN_SUBJECT_SHEET_NOT_FOUND:
      "Appointment Form Submission Error - Sheet Not Found",
    ADMIN_BODY_SHEET_NOT_FOUND:
      'The appointment sheet "{sheetName}" was not found in the linked spreadsheet ({spreadsheetId}). Form details: {submissionDetails}',
    ADMIN_SUBJECT_DUPLICATE: "Appointment Form - Duplicate Submission Detected",
    ADMIN_BODY_DUPLICATE:
      'A duplicate submission for slot "{submittedSlot}" was detected. User: {userEmail}. Form details: {submissionDetails}',
    ADMIN_SUBJECT_NEW_APPOINTMENT: "New Appointment Booked: {submittedSlot}",
    ADMIN_BODY_NEW_APPOINTMENT:
      "A new appointment has been booked.\n\nDetails:\n{submissionDetails}",
    ADMIN_SUBJECT_SUBMISSION_ERROR: "Appointment Form Submission Error",
    ADMIN_BODY_SUBMISSION_ERROR:
      "An error occurred during form submission processing: {errorMessage}\n\nForm details: {submissionDetails}",

    USER_SUBJECT_ERROR: "Error with Your Appointment Submission",
    USER_BODY_ERROR:
      "We experienced an issue processing your appointment. Please contact us directly.",
    USER_SUBJECT_DUPLICATE: "Appointment Not Booked - Slot Taken",
    USER_BODY_DUPLICATE: "{CONFIRMATION_MESSAGES_DUPLICATE_MESSAGE}", // References CONFIRMATION_MESSAGES
    USER_SUBJECT_SUCCESS:
      "Your Dental Appointment is Confirmed for {submittedSlot}!",
    USER_BODY_SUCCESS_PART1:
      "Dear {userName},\n\nThank you for booking your dental appointment with us. Your appointment for {submittedSlot} has been successfully confirmed.\n\nHere are your submission details:\n",
    USER_BODY_SUCCESS_PART2:
      "\n\nWe look forward to seeing you!\n\nBest regards,\nYour Dental Clinic Team",
  },
};
