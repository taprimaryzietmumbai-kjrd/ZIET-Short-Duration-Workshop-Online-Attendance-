# ZIET Mumbai — Workshop Online Attendance

A five-day attendance web application for short-duration workshops conducted by the **Kendriya Vidyalaya Sangathan (KVS), Zonal Institute of Education & Training (ZIET), Mumbai**.

The application uses Google Apps Script for the server-side logic and Google Sheets as the participant source and attendance record.

## Features

- Reads participant lists from worksheet tabs in the configured Google spreadsheet.
- Displays participant names and school details when matching column headers are present.
- Lets a participant select **Present** or **Absent** for each of five workshop days.
- Checks the configured attendance date and time window on the server before saving each day's attendance.
- Requires participants to draw a signature or upload a PNG/JPEG signature image for each session.
- Records the server-side submission timestamp in `Attendance_Log`.
- Places a resized signature image in the corresponding session cell in columns D:M of the selected English or Hindi participant worksheet (Day 1 forenoon/afternoon through Day 5 forenoon/afternoon).
- Includes a responsive interface suitable for desktop and mobile screens.

## Data source

The app is configured to use this spreadsheet:

[ZIET Workshop Participants — Google Sheet](https://docs.google.com/spreadsheets/d/1q7U2KHKFW-DU8ncANE-XXQ-vgdxZDLG2q2uqfxcAuZ4/edit?usp=sharing)

Spreadsheet ID configured in `Code.gs`:

`1q7U2KHKFW-DU8ncANE-XXQ-vgdxZDLG2q2uqfxcAuZ4`

The Google account that deploys/runs the script must have access to this spreadsheet.

## Files

- `Code.gs` — Apps Script backend, sheet discovery, participant retrieval, attendance validation and logging.
- `Index.html` — web interface and calls to the Apps Script backend.
- `README.md` — setup and deployment instructions.

## Setup in Google Apps Script

1. Open [Google Apps Script](https://script.google.com/).
2. Create or open the project for this attendance application.
3. Add the contents of `Code.gs` as the script file named **Code.gs**.
4. Add an HTML file named **Index** (Apps Script will display it as `Index.html`) and paste in the contents of `Index.html`.
5. Confirm the spreadsheet ID in `Code.gs` is the ID shown above.
6. Save the project.
7. Run or open the app and authorize the required Google Sheets access when prompted.
8. Select **Deploy → New deployment → Web app**.
9. Set **Execute as** to the account that has access to the spreadsheet. Set access to the narrowest option suitable for your workshop.
10. Deploy, authorize if prompted, and open the generated **Web app URL**.

When updating an existing deployment, use **Deploy → Manage deployments → Edit → New version → Deploy**.

> The web application must be opened through its Google Apps Script Web app URL. The HTML uses `google.script.run`, which is provided by Apps Script HTML Service and does not work when the page is opened directly from GitHub Pages.

## Configure the five attendance days

On first use, the backend creates an `Attendance_Settings` worksheet if it does not exist. It contains:

| Column | Meaning | Example |
|---|---|---|
| A | Day label | Day 1 |
| B | Date | 2026-10-12 |
| C | Opening time | 09:00 |
| D | Closing time | 17:00 |

Before participants submit attendance:

1. Open the Google Sheet.
2. Open the `Attendance_Settings` tab.
3. Enter the correct workshop date in column B for each of the five days using `YYYY-MM-DD`.
4. Set opening and closing times in columns C and D using 24-hour `HH:MM` format.
5. Check the Apps Script project's time zone in **Project Settings** so it matches the workshop's local time zone.

Blank dates mean attendance for that day cannot be accepted. The backend validates each day separately against the current date and time. Participants should submit during the applicable attendance window.

## Participant worksheet format

Each participant worksheet should have a header row in row 1. Use a recognizable participant-name heading, such as:

- `Participant Name`
- `Name of Participant`
- `Teacher Name`
- `Name`

For school information, use a heading such as:

- `School`
- `School Name`
- `KV Name`
- `Name of KV`

Worksheet tabs intended as participant lists are discovered automatically. The `Attendance_Log` and `Attendance_Settings` tabs are excluded from the list of participant worksheets.

## Attendance records

The backend creates an `Attendance_Log` worksheet when a valid submission is first saved. It stores one row per participant/list/day/session/date, including the server timestamp and the signature cell reference. A later submission for the same participant and session updates the existing log record and replaces the signature image anchored in the corresponding D:M cell. Signature images are inserted as over-grid images anchored to the session cell and resized to 100 × 45 pixels; the target column is set to 110 pixels wide and the participant row to at least 55 pixels high. A note on the cell records status, session and submission time.

## Important security and operational notes

- **Participant identity is not strongly verified in the current version.** Choosing a name from the list does not prove that the person submitting is that participant. Do not treat this alone as secure identity verification.
- Anyone who can access the deployed web app may be able to select another participant unless additional authentication or verification is implemented. Restrict web-app access appropriately.
- Attendance is submitted one session at a time; a participant does not need to mark all days/sessions together.
- Each participant-selected English or Hindi worksheet receives the signature image in columns D:M. Ensure those columns are reserved for attendance signatures, because the script adjusts their width and participant row height to fit the image.
- The `Attendance_Settings` worksheet is configuration data. Limit editing access to workshop organizers.
- Test with a small number of records before using the app for official attendance.
- Do not place passwords, private credentials or OAuth client secrets in this repository.
- GitHub stores the source code. It does not host the Apps Script backend or replace the Google Apps Script deployment.

## Troubleshooting

**“No HTML file named Index was found.”**
- Add an Apps Script HTML file named `Index`. The filename in `doGet()` must match `HtmlService.createHtmlOutputFromFile('Index')`.

**The participant list is empty or names fail to load.**
- Confirm the deploying account can access the spreadsheet.
- Confirm the correct spreadsheet ID is set in `Code.gs`.
- Confirm the participant data is in a worksheet tab and row 1 contains a recognizable name header.

**Attendance is not saved.**
- Check that all five day statuses are selected.
- Confirm each day's date and opening/closing times in `Attendance_Settings`.
- Confirm the Apps Script time zone is correct and the submission is within the permitted window.
- Read the error message shown by the app and check the Apps Script execution log.

**The Google authorization page shows “Error 401: invalid_client.”**
- This is an authorization/OAuth configuration issue, not a participant-sheet error. Check the Apps Script project's Google Cloud project configuration and account permissions. Do not create random OAuth clients or expose credentials.

## Repository

Source code: [ZIET-Short-Duration-Workshop-Online-Attendance-](https://github.com/taprimaryzietmumbai-kjrd/ZIET-Short-Duration-Workshop-Online-Attendance-)

Maintained for workshop attendance workflows at **KVS ZIET Mumbai**.
