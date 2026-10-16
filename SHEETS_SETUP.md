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
- **State**: the full progress JSON the quiz needs to restore a user on another device (answers per question, history, the user's own questions, exam marks). Don't edit it by hand.

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
5. Google asks you to authorize the script. Click **Authorize access** and pick your account. If you see *"Google hasn't verified this app"*, click **Advanced → Go to Quiz Progress API (unsafe)**; this is your own script. Then click **Allow**. The script can only read and write this spreadsheet.
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

If you change `Code.gs`, use **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. This keeps the same URL. *New deployment* would create a new URL that you'd have to paste into `index.html` again.

## How syncing behaves

- **When it saves:** a couple of seconds after each answer, and when the tab is hidden or closed. Every save also goes to the browser's storage, so offline use still works and syncs on the next save.
- **When it loads:** when the quiz opens and whenever the tab comes back into focus, it takes the sheet's copy if that copy is newer.
- **Conflicts:** the newest save wins. The script refuses to overwrite a row with an older copy, so a device that was offline can't wipe newer progress.
- **Usernames:** stored in lowercase. Spaces become dashes, and only letters, numbers, `.`, `_` and `-` are kept.

## Limits to know about

- **No password.** Anyone who knows the Web App URL and a username can read or overwrite that user's progress. That's fine for study progress. Don't store anything sensitive.
- **Cell size.** One cell holds at most 50,000 characters. Normal use is far below that, but hundreds of long custom questions could exceed it. The save is then refused and that user's progress stays in their browser only.
- **Apps Script quotas.** The free quotas cover a study group with plenty of room. Each save is one request, sent at most once every couple of seconds while someone answers questions.
