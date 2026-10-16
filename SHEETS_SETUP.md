# Sync progress with Google Sheets

The quiz can save each user's progress to a Google Sheet so it follows them across devices. There's no login: everyone picks a username the first time they open the quiz, and that username is the row key in the sheet. It takes about 5 minutes to set up and is free.

Until you paste a Web App URL into `index.html`, progress stays in the browser only. The quiz works either way.

## 1. Create the Google Sheet

1. Go to <https://sheets.new> (signed in to the Google account that should own the data).
2. Name it something like **Lakehouse Quiz Progress**.

You don't need to add headers. The script creates a `Progress` tab with these columns on its first request:

| Username | TotalAnswered | TotalCorrect | SectionScores | MockBestPro | MockBestAssoc | LastUpdated | State |
|---|---|---|---|---|---|---|---|

- **SectionScores**: JSON of correct, answered and total per section, for both exams.
- **MockBestPro / MockBestAssoc**: best mock exam percentage for each exam.
- **State**: the full progress JSON the quiz needs to restore a user on another device (answers per question, history, exam marks). Don't edit it by hand.

The first time someone saves a question to their question bank, the script also creates a `Questions` tab with one row per question:

| Username | QuestionId | Exam | SectionId | Section | Question | Code | CorrectAnswer | Option2 | Option3 | Option4 | Explanation | SeenInExam | Updated |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|

Each user's question bank is private to their username. Whenever someone adds, edits or deletes a question, the quiz replaces that user's rows.

Notes get a `Notes` tab the same way, with one row per subheading (or one per heading that has no subheadings yet):

| Username | HeadingId | Heading | SubheadingId | Subheading | Order | Body | Images | Updated |
|---|---|---|---|---|---|---|---|---|

Pictures added to notes are saved in a Google Drive folder called **Lakehouse Quiz Images**, in the account that owns the script. Each picture is shared as "anyone with the link can view" so the quiz can show it. The `Images` column holds the Drive file IDs. Which questions are tagged to which note is part of `State`.

## 2. Paste the Apps Script

1. In the sheet, open **Extensions → Apps Script**.
2. Delete everything in `Code.gs` and paste in the contents of [`apps-script/Code.gs`](apps-script/Code.gs) from this repo.
3. Click **Save** (the disk icon) and give the project a name, e.g. *Quiz Progress API*.

## 3. Deploy it as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear icon next to *Select type* and choose **Web app**.
3. Set:
   - **Description**: `v1`
   - **Execute as**: **Me**
   - **Who has access**: **Anyone**
4. Click **Deploy**.
5. Google asks you to authorize the script. Click **Authorize access** and pick your account. If you see *"Google hasn't verified this app"*, click **Advanced → Go to Quiz Progress API (unsafe)**; this is your own script. Then click **Allow**. The script asks for two permissions: this spreadsheet, and Google Drive (to save pictures added to notes into its own folder).
6. Copy the **Web app URL**. It looks like `https://script.google.com/macros/s/AKfy.../exec`.

## 4. Put the URL in the quiz

Open `index.html` and find this line near the storage code:

```js
var SHEETS_URL="";
```

Paste your URL between the quotes:

```js
var SHEETS_URL="https://script.google.com/macros/s/AKfy.../exec";
```

Save, commit and redeploy the quiz. The footer of the home screen now says progress *syncs to Google Sheets*.

### Check that it works

- Open `YOUR_URL?username=test` in a browser. You should see `{"ok":true,"found":false}`.
- Answer a question in the quiz, wait a couple of seconds, and a row with your username appears in the sheet.
- Open the quiz on another device or browser, enter the same username, and your progress loads.

## Updating the script later

> **Already deployed an earlier version?** The `Questions` and `Notes` tabs, pictures and the group dashboard all need the latest `Code.gs`. Paste it over the old code, save, and redeploy as a new version as described below. Your Web App URL stays the same. Because pictures are stored in Drive, Google asks you to authorize the script again during the redeploy: click **Review permissions** and allow access. Existing questions and notes move into their tabs the next time each user opens the quiz.


If you change `Code.gs`, use **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. This keeps the same URL. *New deployment* would create a new URL that you'd have to paste into `index.html` again.

## How syncing behaves

- **When it saves:** a couple of seconds after each answer, and when the tab is hidden or closed. Every save also goes to the browser's storage, so offline use still works and syncs on the next save.
- **When it loads:** when the quiz opens and whenever the tab comes back into focus, it takes the sheet's copy if that copy is newer.
- **Conflicts:** the newest save wins. The script refuses to overwrite a row with an older copy, so a device that was offline can't wipe newer progress.
- **Usernames:** stored in lowercase. Spaces become dashes, and only letters, numbers, `.`, `_` and `-` are kept.

## Limits to know about

- **No password.** Anyone who knows the Web App URL and a username can read or overwrite that user's progress. That's fine for study progress. Don't store anything sensitive.
- **Cell size.** One cell holds at most 50,000 characters. Questions are stored one per row, so this limit only applies to a single very long field.
- **Editing questions in the sheet.** You can read the `Questions` tab freely. Edits made there are overwritten the next time that user changes their question bank in the quiz, and a device that already has the user's progress may not pick them up.
- **Group dashboard.** Everyone using the quiz can see every username's progress: answers, scores per section, best mock exam and last active time. Notes and question text are never shown there.
- **Pictures.** Anyone with a picture's link can view it, and the upload endpoint accepts images from anyone who has the Web App URL (JPEG, PNG, WebP or GIF, up to about 3 MB each; the quiz shrinks pictures before sending). Pictures removed from a note stay in the Drive folder; delete them there if you want the space back.
- **Apps Script quotas.** The free quotas cover a study group with plenty of room. Each save is one request, sent at most once every couple of seconds while someone answers questions.
