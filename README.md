# Google Apps Script: Dental Appointment Booking System

This Google Apps Script automates a dental appointment booking system using Google Forms and Google Sheets. It manages available time slots, handles new appointments, prevents double-bookgs, and sends automated confirmation emails to both the clinic administrator and the patient.

## ✨ Features

- **Automated Form Creation:** Automatically creates a Google Form with necessary fields if one doesn't exist.
- **Dynamic Time Slots:** Updates available time slots in the Google Form based on a dedicated "Availability" sheet in Google Sheets and existing "Appointment" bookings.
- **Real-time Updates:** Automatically refreshes available slots when the "Availability" sheet is edited or when a new form is submitted.
- **Double-Booking Prevention:** Ensures a time slot can only be booked once.
- **Automated Email Confirmations:** Sends confirmation emails to the clinic administrator and the patient upon successful or failed appointment submissions.
- **Centralized Configuration:** All text strings, email templates, sheet names, and other settings are easily configurable via a single `CONFIG` JSON object.
- **Detailed Logging:** Comprehensive `console.log`, `console.warn`, and `console.error` messages for easy debugging.

---

## 🚀 Getting Started

Follow these detailed steps to set up and deploy your Dental Appointment Booking System.

### Prerequisites

- A Google Account (Gmail, Google Workspace).
- Basic familiarity with Google Forms and Google Sheets.

### Setup Guide

#### Step 1: Create Your Google Form and Link to a Spreadsheet

1.  **Go to Google Forms:** Open your web browser and navigate to [forms.google.com](https://forms.google.com/).
2.  **Start a New Form:** Click on the `+ Blank` option to create a new form. You can give it a temporary title like "My Dental Form".
3.  **Link to a Spreadsheet:**
    - In your new Google Form, click on the **"Responses"** tab.
    - Click the **green Google Sheets icon** (usually on the right side of the "Responses" tab).
    - Select **"Create a new spreadsheet"** and click `CREATE`. This will open a new Google Sheet that is automatically linked to your form. This spreadsheet will initially contain a sheet named "Form Responses 1".

---

#### Step 2: Access the Google Apps Script Editor

1.  **From your Linked Spreadsheet:** With the newly created (or existing) Google Sheet open, go to `Extensions` > `Apps Script`.
2.  This will open a new tab with the Google Apps Script editor. You'll typically see a file named `Code.gs` with some default function.

---

#### Step 3: Copy and Paste the Provided Code

1.  **Delete Existing Code:** In the Apps Script editor, delete any existing code in the `Code.gs` file.
2.  **Paste the New Code:** Copy the entire script provided (or generated previously) and paste it into your empty `Code.gs` file in the Apps Script editor.
3.  **Save the Project:** Click the `Save project` icon (floppy disk) in the toolbar. You can name your project (e.g., "Dental Appointment Script").

---

#### Step 4: Configure the Script (`CONFIG` Object)

The script is designed to be highly configurable through a central `CONFIG` JSON object at the very top of the `Code.gs` file.

1.  **Locate `CONFIG`:** Scroll to the top of your `Code.gs` file and find the `const CONFIG = { ... };` block.
2.  **Update `ADMIN_EMAIL`:**
    - Find the line: `ADMIN_EMAIL: "your.admin.email@example.com",`
    - **Replace `"your.admin.email@example.com"` with your actual email address** where you want to receive administrative notifications (e.g., `clinic.admin@yourdomain.com`). This is crucial for receiving alerts and successful booking notifications.
3.  **Review Other Settings (Optional but Recommended):**
    - **`SHEET_NAMES`**: If you prefer different names for your "Appointment" or "Availability" sheets, you can change them here. The `setup()` function will automatically rename/create them accordingly.
    - **`DEFAULT_SLOT_MESSAGE`**: Customize the message displayed when no time slots are available.
    - **`FORM_TITLES`**: Adjust the titles and descriptions of your form and its fields.
    - **`CONFIRMATION_MESSAGES`**: Modify the messages sent to users and displayed on the form after submission.
    - **`LOG_MESSAGES` / `EMAIL_MESSAGES`**: These contain all the strings used for console logging and email bodies/subjects. You can customize them for localization or specific phrasing.
4.  **Save Changes:** After making your configurations, click the `Save project` icon again.

---

#### Step 5: Prepare Your Spreadsheet Sheets

The script relies on specific sheet names within your linked Google Spreadsheet.

1.  **Rename "Form Responses 1" to "Appointment":**
    - Go back to your Google Sheet.
    - At the bottom, you'll see a sheet tab usually named "Form Responses 1".
    - Right-click on this tab and select `Rename`.
    - Change the name to `Appointment` (case-sensitive, must match `CONFIG.SHEET_NAMES.APPOINTMENTS`).
2.  **Create the "Availability" Sheet:**
    - Click the `+` icon (Add Sheet) at the bottom of your Google Sheet to create a new sheet.
    - Rename this new sheet to `Availability` (case-sensitive, must match `CONFIG.SHEET_NAMES.AVAILABILITY`).
3.  **Populate "Availability" Sheet:**
    - In the `Availability` sheet, enter your available time slots starting from cell `A1`.
    - **Cell A1 should be the header:** You can leave it as "Available Time Slot" or any other descriptive header. The script expects the actual time slots to start from `A2` downwards.
    - **Format for Time Slots:** Enter each time slot on a new row in Column A. The script expects a format like `YYYY-MM-DD HH:MM AM/PM` (e.g., `2025-07-25 09:00 AM - 10:00 AM`). Ensure your dates are in the future for them to be considered valid.
    - **Example `Availability` Sheet:**
      ```
      Available Time Slot
      2025-07-25 09:00 AM - 10:00 AM
      2025-07-25 10:30 AM - 11:30 AM
      2025-07-26 01:00 PM - 02:00 PM
      2025-07-26 02:30 PM - 03:30 PM
      ```

---

#### Step 6: Run the `setup()` Function and Grant Permissions

This is the most critical step to initialize your form and set up the necessary triggers.

1.  **Go back to the Apps Script Editor.**
2.  **Select `setup` function:** In the toolbar, there's a dropdown menu (usually showing "main" or "myFunction"). Click it and select `setup`.
3.  **Run the function:** Click the `Run` button (play icon) next to the dropdown.
4.  **Authorization Required (First Run Only!):**
    - The first time you run `setup()`, Google will prompt you to authorize the script. This is because the script needs permission to access your Google Forms, Google Sheets, and send emails.
    - A dialog box will appear. Click on **`Review permissions`**.
    - You'll be directed to a Google sign-in page. **Select the Google Account** you want to use with this script.
    - You might see a warning message like "Google hasn't verified this app." This is **normal for scripts you write yourself** and are not publicly published.
      - Click on **`Advanced`**.
      - Then click on **`Go to [Project Name] (unsafe)`** (where `[Project Name]` is the name you gave your Apps Script project, e.g., "Go to Dental Appointment Script (unsafe)").
    - Finally, you'll see a screen listing the permissions the script requires (e.g., "View and manage your forms," "View and manage spreadsheets," "Send email as you"). **Review these permissions** and if you're comfortable, click **`Allow`** to grant them.
5.  **Monitor Execution:** After authorization, the script will run automatically. You can view the execution logs by going to `Executions` (on the left sidebar in the Apps Script editor). Look for `Script execution complete` messages, or any error messages.

---

### Step 7: Verify Setup

After the `setup()` function runs successfully:

1.  **Check your Google Form:**
    - Go back to your Google Form (the one you created in Step 1).
    - You should now see the "Name", "Email", "Phone Number", and "Time Slot" fields.
    - The "Time Slot" dropdown should be populated with the available slots you entered in your `Availability` sheet.
2.  **Check Apps Script Triggers:**
    - In the Apps Script editor, click the `Triggers` icon (looks like a clock) on the left sidebar.
    - You should see three new triggers:
      - `onEditAppointmentSheet` (Event type: On edit, Source: From spreadsheet)
      - `onFormSubmitAppointmentSheet` (Event type: On form submit, Source: From form)
      - `onOpenFormHandler` (Event type: On open, Source: From form)

---

## 🚦 Usage

Once setup is complete, your system is ready:

- **Booking an Appointment:** Users fill out and submit the Google Form.
  - Upon submission, the script will record the appointment in the "Appointment" sheet.
  - The booked slot will be removed from the "Time Slot" options in the Google Form.
  - Both the admin (your configured `ADMIN_EMAIL`) and the user (email provided in the form) will receive confirmation emails.
- **Updating Availability:** Edit the "Availability" sheet in your Google Spreadsheet (add new slots, remove old ones). The form's "Time Slot" options will automatically update within a few seconds.
- **Form Opening:** When someone opens the Google Form, the `onOpenFormHandler` trigger will run to ensure the time slots are up-to-date.

---

## 🐛 Troubleshooting

- **"Script doesn't run / Permissions error":** Ensure you've authorized the script correctly in Step 6. Sometimes, re-running `setup()` and re-authorizing can resolve this.
- **"Time slots not updating":**
  - Check the `Availability` sheet name (`Availability`).
  - Ensure the time slots in the `Availability` sheet are in the correct `YYYY-MM-DD HH:MM AM/PM` format and are in the future.
  - Check the Apps Script `Executions` log for any errors in `updateFormSlots`, `fetchAvailabilities`, or `fetchAppointments`.
- **"Emails not sending":**
  - Verify the `ADMIN_EMAIL` in your `CONFIG` is correct.
  - Check the Apps Script `Executions` log for errors in `sendConfirmationEmail`.
  - Ensure your Google Account has permission to send emails (this is part of the initial authorization).
  - Check spam folders for the confirmation emails.
- **"Sheet not found" errors:** Double-check that your sheets are named exactly `Appointment` and `Availability` (case-sensitive).
- **"Lock contention" errors:** These are rare but can occur if many submissions happen simultaneously. The script includes a lock to mitigate this. If persistent, review the `LockService` documentation.
