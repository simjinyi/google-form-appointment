const FORM_NAME = "Appointment Booking Form";
const SLOT_QUESTION_TITLE = "Select Time Slot";
const DEFAULT_CHOICE_VALUE = "No Time Slot Available";

class FormWrapper {
  constructor(spreadsheet, formId = "") {
    if (!spreadsheet) throw new Error(ERROR_DEPENDENCY);

    this.form = getOrCreateForm(formId);
    trySetFormDestination(spreadsheet, this.form);

    this.slotQuestion = trySetupSlotQuestion(this.form);
  }

  updateSlots(slots) {
    const sortedSlots = slots.sort(slotSorterFn);
    this.slotQuestion
      .asListItem()
      .setChoiceValues(
        sortedSlots.length > 0 ? sortedSlots : [DEFAULT_CHOICE_VALUE]
      );
  }

  getForm() {
    return this.form;
  }
}

const getQuestionByTitle = (form, title) => {
  return form.getItems().find((item) => item.getTitle() === title) || null;
};

const getOrCreateForm = (formId = "") => {
  let form = null;

  try {
    form = FormApp.openById(formId);
  } catch (e) {
    form = FormApp.create(FORM_NAME);
  }

  // stop trying after failing for the second time
  if (!form) form = FormApp.create(FORM_NAME);
  if (!form) throw new Error(ERROR_FORM_CREATION);

  return form;
};

const trySetFormDestination = (spreadsheet, form) => {
  // check if destination is the current spreadsheet
  isSpreadsheet =
    form.getDestinationType() === FormApp.DestinationType.SPREADSHEET;
  isIdMatching = form.getDestinationId() === spreadsheet.getId();
  if (isSpreadsheet && isIdMatching) return;

  // caveat: calling this function multiple times would result in multiple sheets getting created
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
};

const trySetupSlotQuestion = (form) => {
  // check if slot question exists in the form
  let slotQuestion = getQuestionByTitle(form, SLOT_QUESTION_TITLE);
  if (slotQuestion) return slotQuestion;

  // if the question failed to be created, do not proceed
  slotQuestion = form.addListItem().setTitle(SLOT_QUESTION_TITLE);
  if (!slotQuestion) throw new Error(ERROR_FIELD_CREATION);

  return slotQuestion;
};

const slotSorterFn = (a, b) => {
  const dateA = new Date(a.replace(" - ", " "));
  const dateB = new Date(b.replace(" - ", " "));

  const isValidA = !isNaN(dateA.getTime());
  const isValidB = !isNaN(dateB.getTime());

  // case 1: both are valid dates
  if (isValidA && isValidB) {
    return dateA.getTime() - dateB.getTime();
  }

  // case 2: only 'a' is a valid date (a comes before b)
  if (isValidA && !isValidB) {
    return -1;
  }

  // case 3: only 'b' is a valid date (b comes before a)
  if (!isValidA && isValidB) {
    return 1;
  }

  // case 4: neither are valid dates (sort alphabetically)
  if (!isValidA && !isValidB) {
    return a.localeCompare(b);
  }

  return 0;
};
