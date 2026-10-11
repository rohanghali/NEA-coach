# Adding the AI key (about 5 minutes, one time, and free)

**The free tier is enough, and you do not need a card.** The AI review costs
nothing on it; what you spend is a daily allowance of REQUESTS, and the app shows
how much of it you have used and when it resets. Measured on a real key, the free
tier allows **20 requests per model per day**, and one full check of a 60-page
portfolio uses about **9** — so roughly **two full checks a day**, more on days when
little has changed, because pages already read are never re-sent.

Because the cap is **per model**, the app uses a cheap model to read the pages and
a stronger one to decide the mark, and shows the usage for each separately. If you
run out on one, the other still has its own allowance.

## What you will see instead of a bill

There is no money anywhere in the app on a free key. The card at the bottom reads
*"Today's free AI allowance"* and shows a percentage used per model, with the time
until it resets. Google resets it at **midnight Pacific time**, which is 8am in the
UK (7am in summer) — the app converts that for you.

If a check is ever refused for using up the allowance, the app says so plainly and
tells you when it comes back, rather than the old "wait a minute", which was wrong
for a daily cap.

---

# The steps

The free checks work without this. The AI review needs a key, because it is the
part that reads your work against the mark scheme and makes judgement calls.

You do **not** need a credit card to start: Google's free tier covers a lot at
this volume.

---

## Step 1 — Get a key

1. Go to **<https://aistudio.google.com/apikey>**
2. Sign in with any Google account.
3. Click **Create API key**.
4. Copy the key. It looks like `AIza...` and is about 39 characters.

If it asks you to pick a project, the default one it offers is fine.

---

## Important: the key is yours, and it is paid for by you

An API key is billed to whoever owns it. If you are passing this app to a friend
or a classmate, **they must make their own key**. The free tier is generous
enough for this app, but it still belongs to one Google account, and that account
is the one that pays if anything is ever charged.

- Never share a key. Do not paste it into a group chat, email it to a friend, or
  leave it visible in a screenshot or a screen recording.
- Your key is not stored inside the `NEA-Coach` folder, so sending someone that
  folder does NOT send your key. It lives in your own user data folder — on a Mac,
  `~/Library/Application Support/NEA Coach/config.json`; on Windows,
  `%APPDATA%\\NEA Coach\\config.json`. Do not send that folder to anyone.
- If a key does leak, open <https://aistudio.google.com/apikey>, **delete that
  key**, and create a new one. Replacing it costs nothing and takes a minute.

---

## Step 2 — Put it in the app (macOS)

1. In the `NEA-Coach` folder, open the `state` folder.
2. Open `config.json` in **TextEdit**.
   If TextEdit reformats it strangely, use right-click → Open With → TextEdit,
   and make sure Format is set to **Plain Text**.
3. Find the line that says `"apiKey": ""` and paste your key between the quotes:

```json
{
  "deckDir": "~/Documents/NEA-Coach/sample",
  "deckFile": null,
  "model": "gemini-3.1-flash-lite",
  "apiKey": "AIzaSy...your key here...",
  "watch": true
}
```

4. Save and close.
5. In the app, press **"Re-check everything"**.

The AI area should stop saying the key is missing. The first run reads every page,
which takes a few minutes for a large portfolio.

---

## Step 2 — Put it in the app (Windows)

1. Open the `NEA-Coach` folder, then the `state` folder inside it.
2. Right-click `config.json` and choose **Open with** → **Notepad**.
   Notepad is the right choice: it saves plain text and never reformats the file.
   Do not use Word (it wants to save a `.docx`) or WordPad, which can quietly
   change the quote characters so the file stops being valid JSON.
3. Find the line that says `"apiKey": ""` and paste your key between the quotes.
   In JSON, a Windows path needs double backslashes:

```json
{
  "deckDir": "C:\\Users\\YourName\\Documents\\NEA-Coach\\sample",
  "deckFile": null,
  "model": "gemini-3.1-flash-lite",
  "apiKey": "AIzaSy...your key here...",
  "watch": true
}
```

4. Save with Control-S and close Notepad.
5. In the app, press **"Re-check everything"**.

The AI area should stop saying the key is missing. The first run reads every page,
which takes a few minutes for a large portfolio.

If Windows hides the file extension, the file is still the one called
`config.json`; make sure you are in `state`, not the top `NEA-Coach` folder.

---

## Step 3 — Check the cost meter

After the first scan, the **cost meter** in the app shows what was spent. Expect
**a few pence** for 200 pages.

You can watch the running total in `<your data folder>/cost.json` at any time.

---

## Keep it safe

- The key lives in `config.json` on this Mac. It is never shown in the app and
  never sent anywhere except to Google, as the password for your own requests.
- **Do not** put this folder in a public place, and do not paste the key into a
  chat, a document, or a screenshot.
- If the key ever leaks, go back to <https://aistudio.google.com/apikey> and
  delete it, then create a new one. It costs nothing to replace.
- Because this folder syncs through Dropbox, the key will be present on any
  machine the folder syncs to. That is fine for your own computers.

---

## If something goes wrong

| Message | What it means | What to do |
|---|---|---|
| "The API key was rejected" | Typo, or a stray space | Copy it again, carefully, with no spaces |
| "Google is rate-limiting this key" | Too many requests too quickly | Wait a minute, press Check now again. Cached pages are not re-charged |
| "does not have access to any Gemini text model" | The key is not enabled for the Gemini API | Create a fresh key at the link above |
| "The model ... was not available, so ... was used instead" | Not an error | Google renamed a model; the app picked a working one automatically |
| `'node' is not recognized`, or "NEA Coach needs Node.js" | Node.js is not installed, or Windows has not picked it up yet | Install the **LTS** from <https://nodejs.org>, then restart the PC if it still complains |
| The launcher window flashes open and vanishes | Something failed before you could read it | Double-click the `.bat` again: it always waits for a key press. If it still vanishes, open a Command Prompt in the folder and run it from there |
| Nothing happens, and the folder path has spaces in it | Nothing is wrong — every path in the launcher is quoted | Keep the app in the folder it came in. Spaces are handled |
| "Windows Defender Firewall has blocked some features" | Windows is asking about the phone view, not about Google | Tick **Private networks only** and allow it. If you never use the phone, Cancel is safe |

---

## What gets sent to Google

Only the **text** extracted from the pages being reviewed, plus your mark scheme.
Your `.pptx` file is never uploaded, and slide images are not sent in this version.

If you would rather nothing left the Mac at all, leave the key blank. The free
checks are fully offline and still catch most mechanical problems.
