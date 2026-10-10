# Putting NEA Coach on GitHub, for free

This gives you one permanent link to send people. They download a zip, unzip it,
double-click one file, and the app runs. **They install nothing** — Node.js is already
inside the zip.

Do this once. It takes about ten minutes.

---

## 1. Make a GitHub account

Go to <https://github.com> and sign up. Free. No card.

---

## 2. Make an empty repository

1. Click the **+** at the top right, then **New repository**.
2. **Repository name:** `nea-coach`
3. **Description:** `Portfolio feedback for GCSE and A-level Design & Technology NEA`
4. Choose **Public**.
   - Public is needed for the free download link to work without people signing in
     as collaborators. The code is yours; you can change this later.
5. **Do NOT** tick "Add a README", "Add .gitignore" or "Choose a licence" — the folder
   already has those, and adding them here causes a conflict on the first push.
6. Click **Create repository**.

Leave that page open. It shows the commands for step 3.

---

## 3. Send the code up

In a terminal, from the app folder:

```bash
cd "/Users/rohanghali/Documents/APPS/DTA Coach/rohan2/NEA-Coach"

git init
git add .
git commit -m "NEA Coach: portfolio feedback for D&T NEA"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/nea-coach.git
git push -u origin main
```

Replace `YOUR-USERNAME` with your GitHub username.

When it asks for a password, **your GitHub password will not work.** You need a
**Personal Access Token** instead:

1. Go to <https://github.com/settings/tokens>
2. **Generate new token** → **Fine-grained token** (or classic, both work)
3. Give it a name like `nea-coach push`, an expiry (90 days is fine), and tick
   **`repo`** (classic) or **Contents: Read and write** (fine-grained).
4. Copy the token and paste it as the password.

**`.gitignore` already protects you here.** It keeps `state/` out of the repository, so
your report, your portfolio copy and **your API key** are never uploaded. Do not remove
that line.

---

## 4. Build the package

```bash
./tools/package.sh windows
```

This writes `dist/NEA-Coach-windows.zip` (about 31 MB), containing the app plus the
Windows copy of Node.js so users install nothing.

---

## 5. Publish it as a release

The zip is too big to upload on GitHub's website (their file picker caps at 25 MB), so
use the command line:

1. Go to <https://github.com/settings/tokens> and make a token with **Contents:
   Read and write** permission on the `nea-coach` repository.
2. Run this, replacing the two placeholders:

```bash
TOKEN=ghp_your_token_here
USER=your-github-username

# Create the release
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Accept: application/vnd.github+json" \
  https://api.github.com/repos/$USER/nea-coach/releases \
  -d '{"tag_name":"v1.0.0","name":"NEA Coach 1.0.0","body":"First release. Unzip, then double-click Start NEA Coach.bat. Get a free Google API key — see SETUP-AI.md."}'

# Upload the zip to it
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/zip" \
  --data-binary @dist/NEA-Coach-windows.zip \
  "https://uploads.github.com/repos/$USER/nea-coach/releases/$(curl -s -H "Authorization: Bearer $TOKEN" https://api.github.com/repos/$USER/nea-coach/releases/latest | grep -m1 '"id"' | tr -dc '0-9')/assets?name=NEA-Coach-windows.zip"
```

If that is too fiddly, the alternative is to install the GitHub CLI once
(`brew install gh`, then `gh auth login`) and run:

```bash
gh release create v1.0.0 dist/NEA-Coach-windows.zip --title "NEA Coach 1.0.0" --notes "Unzip, then double-click Start NEA Coach.bat."
```

---

## 6. The link to send people

Once the release exists, your download link is:

```
https://github.com/YOUR-USERNAME/nea-coach/releases/latest
```

That page always points at the newest version. You can put that in a message, a
WhatsApp group, or a school email.

To publish an update later:

```bash
./tools/package.sh windows
# then repeat step 5 with a new tag, e.g. v1.0.1
```

---

## What the person receiving it does

1. Click the link, download `NEA-Coach-windows.zip`.
2. Right-click the zip → **Extract All**.
3. Open the folder and double-click **`Start NEA Coach.bat`**.
4. **Leave the black window open** — that is the app running.
5. Get a free Google API key (Settings has a **Paste a key** box, and `SETUP-AI.md`
   explains it). Free, no card.

If Windows shows a blue "Windows protected your PC" box, that is SmartScreen being
cautious about a new, unsigned program. Click **More info** → **Run anyway**.

---

## Before you tell anyone about it

- **The `.gitignore` must keep `state/` out.** Check with `git status` before your
  first push: you should see no `state/` files listed. If you ever see them, stop and
  remove them before pushing.
- **Never commit a release zip** — it is in `.gitignore` for that reason.
- **Do not put your API key in any file that gets uploaded.** The app reads it from
  `state/config.json`, which is ignored.

---

## A note on selling it later

GitHub gives you a free download link; it cannot take payment. If you decide to charge
for it, the payment part is a separate tool (Stripe, Gumroad or Lemon Squeezy).

Two things to know before that:

1. **Each buyer must use their own Google API key.** Google's terms do not allow
   reselling API access, and if everyone used yours you would pay for every check they
   ran.
2. **The app never sends anyone's work to you.** That is its best feature and worth
   keeping, because it means you are not storing other people's coursework.
