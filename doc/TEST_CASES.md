# Teacher Assistant – PowerPoint Add-in
## Test Cases for Microsoft AppSource Certification

**App name:** Teacher Assistant  
**Version:** 1.0.0.0  
**Backend:** https://teachers-center-be.onrender.com  
**Frontend:** https://teachers-center-powerpoint.onrender.com  
**Tester note:** The backend is hosted on Render free tier. The first request after a period of inactivity may take up to 30 seconds to respond while the server wakes up. Subsequent requests respond in 3–8 seconds.

---

## Environment

- PowerPoint for Windows (Desktop), version 16.0 or later
- Internet connection required (AI features call OpenAI API via backend)
- No account or login required

---

## TC-01 — First Launch: Settings Modal Appears

**Category:** Onboarding  
**Priority:** Critical

**Steps:**
1. Open Microsoft PowerPoint
2. Go to the Home tab in the ribbon
3. Click the "Teacher Assistant" button
4. The taskpane sidebar opens on the right side

**Expected result:**
- The settings modal opens automatically on first launch
- Modal contains: Language dropdown, Level dropdown, Age Group field, Cancel and Save buttons
- The input area is blocked (user cannot type) until settings are confirmed
- Clicking outside the modal does nothing (modal stays open on first launch)

---

## TC-02 — Settings: Save and Persist

**Category:** Settings  
**Priority:** Critical

**Steps:**
1. In the settings modal, set Language to "Spanish"
2. Set Level to "A1"
3. Leave Age Group empty
4. Click "Save"

**Expected result:**
- Modal closes
- Context badge in the header updates to "A1 Spanish"
- Input field becomes active and accepts text
- Settings are saved to localStorage for this file

**Persistence check:**
1. Close the taskpane and reopen it
2. Click the context badge (top-right of taskpane)

**Expected result:**
- Settings modal shows previously saved values (Spanish / A1)

---

## TC-03 — Settings: Update Language and Level Mid-Session

**Category:** Settings  
**Priority:** High

**Steps:**
1. Confirm initial settings (any language/level)
2. Click the context badge (e.g. "B1 English") in the taskpane header
3. Change Language to "German" and Level to "B2"
4. Click "Save"

**Expected result:**
- Modal closes
- Context badge updates to "B2 German"
- Next generation request uses the new settings

---

## TC-04 — Settings: Cancel Closes Modal Only if Already Confirmed

**Category:** Settings  
**Priority:** Medium

**Steps (first launch):**
1. Open taskpane for the first time
2. Click "Cancel" in the settings modal

**Expected result:**
- Modal stays open (cannot be dismissed before first save)

**Steps (after settings saved):**
1. Click the context badge to open settings
2. Click "Cancel"

**Expected result:**
- Modal closes without saving changes
- Context badge still shows previous settings

---

## TC-05 — Generate Vocabulary Slides

**Category:** Content Generation  
**Priority:** Critical

**Pre-condition:** Settings confirmed (Language: Spanish, Level: A1)

**Steps:**
1. Select "Vocabulary" from the type selector buttons
2. Type: "Give me 5 food words"
3. Press Enter or click the send button

**Expected result:**
- User message appears in the chat
- Progress indicator shows: "Generating content..." → "Thinking..." → "Creating..." → "Polishing..."
- Slide preview appears with a title slide + 5 content slides (6 total)
- First slide shown: Title slide with presentation title
- Counter shows "Slide 1 of 6"
- Navigation buttons visible: Back (disabled), Skip, Edit, Next

---

## TC-06 — Generate Grammar Slides

**Category:** Content Generation  
**Priority:** Critical

**Pre-condition:** Settings confirmed (Language: English, Level: B1)

**Steps:**
1. Select "Grammar" from the type selector
2. Type: "Explain the present perfect tense with 3 examples"
3. Press Enter

**Expected result:**
- Slides generated covering the grammar rule with examples
- Title slide + content slides with explanations
- Preview navigable with Back/Next buttons

---

## TC-07 — Generate Quiz Slides

**Category:** Content Generation  
**Priority:** Critical

**Pre-condition:** Settings confirmed (Language: English, Level: B2)

**Steps:**
1. Select "Quiz" from the type selector
2. Type: "Create 4 multiple choice questions about travel vocabulary"
3. Press Enter

**Expected result:**
- Quiz slides generated with questions and answer options (A, B, C, D)
- Title slide + question slides
- Content is appropriate for B2 level

---

## TC-08 — Generate Homework Slides

**Category:** Content Generation  
**Priority:** Critical

**Pre-condition:** Settings confirmed (Language: English, Level: A2)

**Steps:**
1. Select "Homework" from the type selector
2. Type: "Create a fill-in-the-blank exercise about daily routines"
3. Press Enter

**Expected result:**
- Homework slides generated with tasks/exercises
- Content appropriate for A2 learners

---

## TC-09 — AI Asks Clarifying Question (Requirements Not Met)

**Category:** Content Generation  
**Priority:** High

**Steps:**
1. Select "Vocabulary" from the type selector
2. Type a vague request: "Give me some words"
3. Press Enter

**Expected result:**
- AI responds with a clarifying question (e.g. "How many words would you like?")
- No slides are generated
- User can answer the question and continue the conversation

**Follow-up steps:**
1. Type: "Give me 5 words about animals"
2. Press Enter

**Expected result:**
- Slides are generated based on the combined context

---

## TC-10 — Slide Preview Navigation with Buttons

**Category:** Preview Navigation  
**Priority:** Critical

**Pre-condition:** Slides generated (at least 3)

**Steps:**
1. Verify Back button is disabled on slide 1
2. Click "Next" → advances to slide 2
3. Click "Next" → advances to slide 3
4. Click "Back" → returns to slide 2
5. Navigate to the last slide

**Expected result:**
- Slide counter updates correctly ("Slide X of Y")
- Back button disabled only on slide 1
- On the last slide, Next button label changes to "Insert N" (e.g. "Insert 6")

---

## TC-11 — Slide Preview Navigation with Keyboard

**Category:** Keyboard Shortcuts  
**Priority:** High

**Pre-condition:** Slides generated, input field NOT focused

**Steps:**
1. Press → (right arrow) → advances to next slide
2. Press ← (left arrow) → goes to previous slide
3. Press Backspace → goes to previous slide
4. Press Enter (when input is empty) → advances to next slide
5. On last slide, press Enter → inserts all slides

**Expected result:**
- Each key triggers the corresponding navigation action
- Visual button flash confirms the key press

---

## TC-12 — Skip (Remove) a Slide

**Category:** Preview Navigation  
**Priority:** High

**Pre-condition:** 6 slides generated

**Steps:**
1. Navigate to slide 3
2. Click "Skip" button (or press R)

**Expected result:**
- Slide 3 is removed
- Counter updates to "Slide 3 of 5" (stays on same index)
- Remaining slides renumber correctly

**Edge case — remove last remaining slide:**
1. Skip all slides one by one until 1 remains
2. Click Skip

**Expected result:**
- Preview dismissed with message "All slides removed."

---

## TC-13 — Edit a Slide

**Category:** Edit Mode  
**Priority:** Critical

**Pre-condition:** Slides generated

**Steps:**
1. Navigate to slide 2
2. Click "Edit" button (or press E)

**Expected result:**
- Edit badge appears: "Editing Slide 2"
- Input placeholder changes to "Describe what to change on slide 2..."
- Type selector is hidden
- Input field is focused

**Steps (continue):**
1. Type: "Add a more detailed explanation"
2. Press Enter

**Expected result:**
- Progress shows "Updating slide..."
- Slide 2 content updates in place
- Preview returns to slide 2 with updated content
- AI message "Slide updated." appears in chat

---

## TC-14 — Edit Mode: Navigate Between Slides While Editing

**Category:** Edit Mode  
**Priority:** Medium

**Pre-condition:** Edit mode active on slide 2

**Steps:**
1. Click "Back" to go to slide 1
2. Click "Next" to go to slide 3

**Expected result:**
- Edit badge updates: "Editing Slide 1" then "Editing Slide 3"
- Input placeholder updates to match current slide number
- Edit instruction will apply to whichever slide is currently shown

---

## TC-15 — Exit Edit Mode with Escape Key

**Category:** Edit Mode  
**Priority:** Medium

**Pre-condition:** Edit mode active

**Steps:**
1. Press Escape

**Expected result:**
- Edit badge disappears
- Input placeholder resets to "Type your request..."
- Type selector reappears
- Edit mode exited without submitting anything

---

## TC-16 — Exit Edit Mode with X Button

**Category:** Edit Mode  
**Priority:** Medium

**Pre-condition:** Edit mode active

**Steps:**
1. Click the X button on the edit badge

**Expected result:**
- Same as TC-15: edit mode exited cleanly

---

## TC-17 — Insert All Slides into PowerPoint

**Category:** Slide Insertion  
**Priority:** Critical

**Pre-condition:** 6 slides generated and previewed

**Steps:**
1. Navigate to the last slide
2. Click "Insert 6" button (or press A from any slide)

**Expected result:**
- Progress shows "Inserting slides..." → "Inserting slide X of 6..."
- Slides are added to the active PowerPoint presentation
- Title slides have: centered title (red, bold, 44pt), subtitle (gray, 24pt)
- Content slides have: title (red, bold, 32pt), content text (black, 18pt), example (italic, gray, 16pt)
- Success message: "6 slides inserted successfully"
- Chat and conversation state are cleared after insertion
- Welcome state resets

---

## TC-18 — Insert All Slides with Keyboard Shortcut A

**Category:** Keyboard Shortcuts  
**Priority:** Medium

**Pre-condition:** Slides in preview, input NOT focused

**Steps:**
1. Press A

**Expected result:**
- All slides inserted immediately (same as clicking Insert button)

---

## TC-19 — Cancel Generation During Progress

**Category:** Cancel  
**Priority:** High

**Steps:**
1. Send a request and wait for progress to start
2. Click the Cancel button in the progress indicator (or press Q)

**Expected result:**
- Progress indicator disappears
- AI message: "Generation cancelled."
- Input becomes active again
- Any response arriving after cancellation is ignored (stale response handling)

---

## TC-20 — Cancel Preview

**Category:** Cancel  
**Priority:** High

**Pre-condition:** Slides in preview

**Steps:**
1. Press Q (while input is NOT focused)

**Expected result:**
- Preview dismissed
- AI message: "Preview cancelled."
- Input active, ready for new request

---

## TC-21 — New Request While Preview is Active

**Category:** State Management  
**Priority:** High

**Pre-condition:** Slides in preview (e.g. 6 slides)

**Steps:**
1. Type a new request in the input field
2. Press Enter

**Expected result:**
- Existing preview is dismissed with message "6 slides not inserted"
- New request is sent
- New slides generated and shown in preview

---

## TC-22 — New Chat Button

**Category:** State Management  
**Priority:** High

**Steps:**
1. Have a conversation with multiple messages and a preview
2. Click the "New Chat" button (top of taskpane)

**Expected result:**
- All messages cleared from chat
- Preview dismissed
- Conversation ID reset
- Welcome state shown again
- Settings are NOT reset (language/level preserved)

---

## TC-23 — Multi-Turn Conversation (Follow-up Request)

**Category:** Conversation Context  
**Priority:** High

**Steps:**
1. Generate vocabulary slides: "Give me 5 animal words"
2. Insert the slides
3. Type: "Now give me 5 more animal words but different ones"
4. Press Enter

**Expected result:**
- AI uses conversation history context
- New set of 5 different animal words generated
- No repetition of words from the first request

---

## TC-24 — Settings Saved Per File

**Category:** Settings Persistence  
**Priority:** Medium

**Steps:**
1. Open File A → set Language to "Spanish", Level to "A1" → Save
2. Open File B → set Language to "French", Level to "C1" → Save
3. Switch back to File A and reopen taskpane

**Expected result:**
- File A shows "A1 Spanish"
- File B shows "C1 French"
- Settings are independent per file

---

## TC-25 — Backend Connection Error

**Category:** Error Handling  
**Priority:** High

**Steps:**
1. Disconnect from the internet
2. Type a request and press Enter

**Expected result:**
- Error message: "Not connected to server. Make sure the backend is running."
- Input becomes active again

---

## TC-26 — Backend Reconnection (Auto-Retry)

**Category:** Error Handling  
**Priority:** Medium

**Steps:**
1. Disconnect from the internet while taskpane is open
2. Wait 10 seconds
3. Reconnect to the internet
4. Send a new request

**Expected result:**
- WebSocket reconnects automatically (up to 3 attempts, 2s backoff)
- Request sends successfully after reconnection

---

## TC-27 — OpenAI API Quota Error

**Category:** Error Handling  
**Priority:** Medium

**Simulated by:** backend returning quota error

**Expected result:**
- Error message displayed to user (user-friendly, not a raw API error)
- Input available for retry

---

## TC-28 — Empty Input Blocked

**Category:** Input Validation  
**Priority:** Medium

**Steps:**
1. Leave the input field empty
2. Press Enter

**Expected result:**
- Nothing happens; no request sent

---

## TC-29 — Input Blocked While Processing

**Category:** Input Validation  
**Priority:** Medium

**Steps:**
1. Send a request
2. While the progress indicator is showing, try typing in the input

**Expected result:**
- Input is disabled during processing
- Placeholder shows "Generating content..."

---

## TC-30 — Keyboard Shortcuts Inactive While Input is Focused

**Category:** Keyboard Shortcuts  
**Priority:** Medium

**Pre-condition:** Slides in preview, input field focused

**Steps:**
1. Press R, E, A, Q

**Expected result:**
- No navigation actions triggered (keys type into the input field normally)
- Shortcuts only active when input is NOT focused

---

## TC-31 — Long Request Input Auto-Resizes

**Category:** UI  
**Priority:** Low

**Steps:**
1. Type a very long multi-line message (paste several sentences)

**Expected result:**
- Textarea grows in height up to max ~100px
- Does not overflow the input container

---

## TC-32 — Single Slide Generation and Insert

**Category:** Edge Cases  
**Priority:** Medium

**Steps:**
1. Type: "Give me 1 word for apple in Spanish"
2. Press Enter

**Expected result:**
- Preview shows 1 or 2 slides (title + content)
- Insert button shows "Insert 1" or "Insert 2"
- Slides insert correctly

---

## TC-33 — Title Slide Formatting After Insert

**Category:** Slide Formatting  
**Priority:** High

**Steps:**
1. Generate any slides and insert them
2. Open the inserted slides in PowerPoint and inspect visually

**Expected result:**
- Title slide: title text centered, red (#d13438), bold, 44pt
- Title slide: subtitle text centered, gray (#605e5c), 24pt
- Content slides: title red, bold, 32pt (top-left)
- Content slides: body text black (#323130), 18pt
- Content slides: example text italic, gray (#605e5c), 16pt
- No leftover default placeholder shapes

---

## TC-34 — Age Group Setting Sent to Backend

**Category:** Settings  
**Priority:** Low

**Steps:**
1. Open settings, set Age Group to "12-15"
2. Save settings
3. Generate vocabulary slides

**Expected result:**
- Generated content is age-appropriate (verified by reading slide content)
- Age group value is included in the backend request

---

## Summary Checklist

| TC | Description | Priority |
|---|---|---|
| TC-01 | First launch settings modal | Critical |
| TC-02 | Settings save and persist | Critical |
| TC-03 | Update settings mid-session | High |
| TC-04 | Cancel settings behavior | Medium |
| TC-05 | Generate vocabulary slides | Critical |
| TC-06 | Generate grammar slides | Critical |
| TC-07 | Generate quiz slides | Critical |
| TC-08 | Generate homework slides | Critical |
| TC-09 | AI clarifying question flow | High |
| TC-10 | Preview navigation (buttons) | Critical |
| TC-11 | Preview navigation (keyboard) | High |
| TC-12 | Skip/remove slide | High |
| TC-13 | Edit a slide | Critical |
| TC-14 | Navigate slides while in edit mode | Medium |
| TC-15 | Exit edit mode with Escape | Medium |
| TC-16 | Exit edit mode with X button | Medium |
| TC-17 | Insert all slides | Critical |
| TC-18 | Insert with A shortcut | Medium |
| TC-19 | Cancel generation | High |
| TC-20 | Cancel preview | High |
| TC-21 | New request dismisses old preview | High |
| TC-22 | New chat button | High |
| TC-23 | Multi-turn conversation context | High |
| TC-24 | Settings saved per file | Medium |
| TC-25 | Backend connection error | High |
| TC-26 | Auto-reconnect after disconnect | Medium |
| TC-27 | OpenAI quota error handling | Medium |
| TC-28 | Empty input blocked | Medium |
| TC-29 | Input blocked during processing | Medium |
| TC-30 | Shortcuts inactive while typing | Medium |
| TC-31 | Long input auto-resize | Low |
| TC-32 | Single slide generation | Medium |
| TC-33 | Slide formatting after insert | High |
| TC-34 | Age group setting used | Low |
