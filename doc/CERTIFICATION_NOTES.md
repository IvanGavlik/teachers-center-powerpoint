# Teacher Assistant – Certification Notes for Microsoft Reviewers

Version 2.0.0.0

No account or login is required to use the add-in.

**Note on first use:** our server may need up to a minute to wake up if it has not been used
recently. If the first request shows a connection message, wait a moment and send it again.

**Sample document for testing "Attach a book":**
https://teacher-assistant.center/assets/sample-book.pdf — please download it before you start.

---

## What's new in this version

- **Attach a book** — attach a PDF (for example a coursebook unit) and slides, edits and games
  are based on its content.
- **New slide design** — slides are inserted as structured layouts (bullet lists, grammar
  patterns, comparisons, vocabulary tables, example sentences) that fit the slide size.
- **What's new dialog** — explains the new features on first start.

---

## How to Install and Open the Add-in

1. Open Microsoft PowerPoint (Windows desktop, version 16.0 or later; PowerPoint on the web also works)
2. Open a new blank presentation
3. Go to the **Home** tab in the ribbon
4. Click the **"Teacher Assistant"** button
5. The taskpane opens on the right side of the screen

---

## Step 1 — Settings and What's New (first start)

- On first launch, a settings dialog opens automatically
- Set Language to **English**, Level to **B1**, leave Age Group empty
- Click **Save** — the header badge shows "B1 English"
- The **"Get familiar with new improvements"** dialog opens. It has two pages:
  - Page 1: **Attach a Book**
  - Page 2: **Interactivity Mode**
- Move between pages with **Next** / **Back**, the dots, or the ← / → keys
- Click **Got it** on the last page to close it (tick "Don't show this again" to stop it
  appearing on the next start)

## Step 2 — Generate slides

- Type: `Explain the present simple for daily habits in 3 slides`
- Press **Enter** or click the send button
- Progress messages appear while the AI works (usually 10–40 seconds)
- A preview appears in the taskpane: a title card followed by the content slides

## Step 3 — Review the preview

- **Next** (→ / Enter) and **Back** (Backspace) move between slides
- **Remove** (bin icon / R) removes the current slide from the preview
- **Edit** (pencil icon / E) — type an instruction such as `Make it simpler` and press Enter;
  the slide updates in place. Press Escape to leave edit mode.

## Step 4 — Insert into PowerPoint

- On the last slide the Next button becomes **"Insert N"** (N = number of slides)
- Click it or press Enter — the slides are added after the last slide of the presentation, in
  the presentation's own theme
- A success message confirms the insert

## Step 5 — Attach a book (new)

1. Click the **paperclip** button in the message box and select the downloaded
   `sample-book.pdf` (PDF only, up to 20 MB)
2. A chip above the message box shows **Uploading… → Processing… → Ready** (usually under a
   minute). Sending is blocked until it is ready.
3. Type: `Make 3 slides with the vocabulary from Unit 3` and press Enter
4. Your message shows "📄 Using: sample-book.pdf". The slides use the vocabulary from the book
   (booking, specials, starter, bill, tip …). Asking `Make slides about the reading text`
   produces slides about **The Copper Kettle**, **Marta** and **Leo** — names that appear only
   in the sample document, which shows the book was used.
5. Click **✕** on the chip to remove the book. Because the conversation was based on the book,
   the add-in asks for confirmation; confirming starts a new chat.

## Step 6 — Interactive games

1. Click **New Chat** (top of the taskpane)
2. Type `/` in the message box — a list of four games appears: multiple choice, sentence
   ordering, guess the word, true or false
3. Select **/multiple-choice** (a green chip shows the mode), type `past simple of irregular verbs`
   and press Enter
   - Alternatively type `make a quiz on past simple` — the add-in asks "Make it interactive" or
     "Create slides"
4. The questions appear in the preview; move to the last card and click **Insert**
5. One slide is inserted with a **QR code** and "Scan to play!". Scanning it with a phone (or
   opening the link in a browser) opens the game. Students do not need an app or an account,
   and no personal data is collected.

## Other things to check

- **Cancel** — while progress is showing, click ✕ or press Q
- **New chat** — the New Chat button resets the conversation
- **Clarifying questions** — a vague request such as `teach something` gets a question back
  instead of slides
- **Settings** — click the "B1 English" badge to change language, level or age group
- **Feedback** — the feedback button in the header, and the thumbs up / down after each answer

---

## Keyboard Shortcuts Reference

The letter and arrow shortcuts work while a preview is open and the cursor is not in the
message box (click anywhere on the preview first).

| Key | Action |
|---|---|
| Enter | Send message / next slide / insert on the last slide |
| → | Next slide |
| ← / Backspace | Previous slide |
| E | Edit the current slide |
| R | Remove the current slide |
| A | Insert all slides |
| Q | Cancel generation or close the preview |
| Escape | Leave edit mode |

---

## Data and privacy

- No account, sign-in or payment
- Requests and uploaded PDFs are processed by OpenAI to generate content. Uploaded PDFs are
  deleted automatically (the file after 7 days, its index 3 days after last use)
- Privacy policy: https://teacher-assistant.center/privacy-policy.html
- Support: https://teacher-assistant.center/support.html
