Here's a full manual testing guide. Use it through the PowerPoint add-in (dev mode) or any WebSocket client like
  Postman/wscat pointed at ws://localhost:2000/ws.

  ---
  Settings to use for most tests

  Language: English
  Level: B1
  Age Group: 14–16
  Native Language: No

  ---
  PATH 1 — Validation / Missing Info

  These should never generate slides. GPT should reply with a clarification question.

  ---
  TC-01 — No topic, no slide count
  I need some vocabulary slides

  Expected: GPT asks "What topic would you like?" and/or "How many slides do you need?"
  
  OK

  ---
  TC-02 — Topic given but no slide count
  I need vocabulary slides about animals

  Expected: GPT asks for the number of slides
  
  OK

  ---
  TC-03 — Slide count given but no topic
  Give me 5 slides

  Expected: GPT asks what topic/content the slides should be about
  
  OK

  ---
  TC-04 — Completely vague
  Help me make a lesson

  Expected: GPT asks what type of content (explanation, task, quiz?), what topic, how many slides

  OK
  ---
  TC-05 — Wrong content implied (no explicit type)
  I want something about Present Simple

  Expected: GPT asks — do you want explanation, examples, tasks, or a quiz on Present Simple?
  
  OK	
  ---
  PATH 2 — Successful Generation

  These should produce slides JSON.

  ---
  TC-06 — Clean vocabulary request
  I need 3 vocabulary slides about food and restaurants

  Expected:
  - 3 slides
  - Each slide has a title and content
  - Content uses symbols (✓ ✗ →), short lines, no paragraphs
  - No translation field (native language = No)

	NOK

  ---
  TC-07 — Grammar explanation
  I need 2 slides explaining Present Simple affirmative and negative forms

  Expected:
  - 2 slides, one for affirmative, one for negative
  - Content uses transformation arrows (→), pattern formatting
  - Rule + example format, no mixing of too many concepts per slide
	
	OK 	

  ---
  TC-08 — Task/exercise type (not explanation)
  I need 3 task slides on Present Simple, use food vocabulary in the tasks

  Expected:
  - 3 slides with tasks/exercises
  - NOT explanations or definitions — actual student tasks
  - Food-related sentences
	
	OK
 
 ---
  TC-09 — Quiz
  I need 2 quiz slides about travel vocabulary

  Expected: Quiz-format slides with questions or fill-in-the-blank style content

  ---
  TC-10 — Homework
  I need 1 homework slide on irregular verbs

  Expected: A homework assignment slide, not an explanation

  ---
  TC-11 — Many slides (cognitive load check)
  I need 6 slides on Present Simple — use, form, third person rule, negatives, questions, short answers

  Expected:
  - 6 separate slides, one concept each
  - NOT one slide overloaded with all of it
  - Max ~5–7 lines per slide

  ---
  PATH 3 — Language & Translation

  ---
  TC-12 — Request in German (teacher writes in German)

  Settings: Language = German, Level = A2, Age Group = 10–12
  Ich brauche 3 Folien über Tiere

  Expected: Clarification OR slides — both should be in German, not English
  
  OK

  ---
  TC-13 — Native language enabled

  Settings: Language = English, Level = A1, Native Language = German
  I need 3 vocabulary slides about colours

  Expected:
  - Slides generated
  - Each slide has a translation field with German translations

  NOK

  ---
  TC-14 — Native language = No (explicit check)

  Settings: Native Language = No
  I need 2 vocabulary slides about weather

  Expected: No translation field anywhere in the response
  
  OK

  ---
  PATH 4 — Conversation Flow (multi-turn)

  These test that GPT remembers context from earlier in the chat.

  ---
  TC-15 — GPT asks, teacher answers, content generates

  Turn 1: I need vocabulary slides about sports
  → Expected: GPT asks how many slides

  Turn 2: 3 slides please
  → Expected: 3 slides about sports vocabulary generated
	
	OK 
  ---
  TC-16 — Start new chat resets context

  After TC-15, click "New Chat" then:
  Give me 3 slides please

  Expected: GPT asks for topic again — it does NOT remember the previous "sports" context

  OK	

  ---
  PATH 5 — Slide Editing

  These require slides to already be generated and visible in preview.

  ---
  TC-17 — Edit content of one slide

  First generate slides (e.g. TC-06).
  Then click edit on slide 1 and type:
  Make the content shorter, maximum 4 lines

  Expected: That single slide returns with condensed content, other slides unchanged
  
  OK

  ---
  TC-18 — Change a slide's topic

  Generate slides, edit slide 2:
  Change the topic of this slide to drinks instead of food

  Expected: Slide 2 regenerated with drinks vocabulary, same format
  
  OK

  ---
  TC-19 — Add examples to a slide

  Generate a grammar slide, edit it:
  Add 2 example sentences showing the rule in action

  Expected: Same slide returned with 2 new example sentences added
  
  OK

  ---
  TC-20 — Make slide simpler (level adjustment)

  Edit a slide:
  Simplify this for A1 students, use very basic vocabulary

  Expected: Simpler language, shorter words, no complex structures
  
  OK

  ---
  PATH 6 — Edge Cases & Stress

  ---
  TC-21 — Request irrelevant to language teaching
  Can you write me a business email?

  Expected: GPT should either refuse politely or ask to clarify in the context of language teaching (the system prompt
  constrains it to language teacher role)
	
  OK
  ---
  TC-22 — Request violates level appropriateness

  Settings: Level = A1
  I need 3 slides on subjunctive mood with idiomatic expressions

  Expected: Either GPT generates appropriately simplified content for A1, or it flags that subjunctive is beyond A1
  level and asks to confirm

  ---
  TC-23 — Ambiguous slide count
  I need a few slides about jobs vocabulary

  Expected: GPT asks for an exact number — the prompt says "Do NOT assume defaults"

  ---
  TC-24 — Request asks for something half-explained
  I need slides about verbs

  Expected: GPT asks — which verbs? what aspect (form, use, conjugation)? how many slides?

  ---
  What to watch for in every response


TODO FIX
  


---

deleta all slideds what happens ? test this 
- all slides remove what about chat histora 

todo ask to make test for chat history double check the code see where at which palces and when is should be delete it is not only on create 
----

TEST EDIT ask to change layout of the slide (look) not the content 





 ---t

  ---
  1. First Launch — Settings Gate

  Goal: Confirm settings must be configured before anything can be typed.

  ┌──────┬─────────────────────────────────────────────┬────────────────────────────────────────────────────────────┐
  │ Step │                   Action                    │                      Expected Result                       │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.1  │ Open PowerPoint with the add-in sideloaded  │ Sidebar loads, welcome screen shows "What would you like   │
  │      │                                             │ to create?"                                                │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.2  │ Click the message input box                 │ Settings modal opens automatically (input is blocked until │
  │      │                                             │  settings confirmed)                                       │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.3  │ Click outside the modal (backdrop)          │ Modal closes without saving                                │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.4  │ Click the context badge in the header       │ Modal opens again                                          │
  │      │ (shows "B1 English")                        │                                                            │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.5  │ Change Language to German, Level to A2,     │ Modal closes, badge updates to "A2 German", input is now   │
  │      │ click Save                                  │ usable                                                     │
  ├──────┼─────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 1.6  │ Click badge again, click Cancel             │ Settings revert to previously saved values                 │
  └──────┴─────────────────────────────────────────────┴────────────────────────────────────────────────────────────┘

  ---
  2. Settings — All Combinations

  Goal: Confirm all settings options are accepted and reflected in the badge.

  ┌──────┬──────────────────────────────────┬─────────────────────────────────────────────────────────────────────┐
  │ Step │              Action              │                           Expected Result                           │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.1  │ Set Language: English, Level: A1 │ Badge: "A1 English"                                                 │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.2  │ Set Language: French, Level: B2  │ Badge: "B2 French"                                                  │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.3  │ Set Language: Spanish, Level: C1 │ Badge: "C1 Spanish"                                                 │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.4  │ Set Language: Italian, Level: C2 │ Badge: "C2 Italian"                                                 │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.5  │ Set Native Language: German      │ Native language column should appear in generated vocabulary slides │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.6  │ Set Native Language: No          │ No translation column in slides                                     │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.7  │ Set Age Group: "12-14"           │ Content should reflect younger audience                             │
  ├──────┼──────────────────────────────────┼─────────────────────────────────────────────────────────────────────┤
  │ 2.8  │ Leave Age Group empty            │ Content generated without age constraint                            │
  └──────┴──────────────────────────────────┴─────────────────────────────────────────────────────────────────────┘

  ---
  3. Vocabulary — Happy Path

  Goal: Full generate → preview → insert flow for vocabulary.

  Settings: English, B1, Native Language: No

  ┌──────┬──────────────────────────────────────┬───────────────────────────────────────────────────────────────────┐
  │ Step │                Action                │                          Expected Result                          │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.1  │ Type: Give me 5 food vocabulary      │ Progress spinner appears with stages: Generating content… →       │
  │      │ words with examples                  │ (backend stages) → preview appears                                │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.2  │ Check preview header                 │ Shows "Preview — N slides", "Vocabulary" badge in corner          │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.3  │ Check slide 1 (Title slide)          │ Shows topic title and subtitle with level                         │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.4  │ Click Next (or press Enter)          │ Advances to slide 2 (first word)                                  │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.5  │ Verify slide card                    │ Has word as title, definition as content, example sentence        │
  │      │                                      │ visible                                                           │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.6  │ Press Enter repeatedly through all   │ Navigates forward each time                                       │
  │      │ slides                               │                                                                   │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.7  │ On last slide, check Next button     │ Should read "Insert N" instead of "Next"                          │
  │      │ label                                │                                                                   │
  ├──────┼──────────────────────────────────────┼───────────────────────────────────────────────────────────────────┤
  │ 3.8  │ Click Insert N                       │ Slides appear in PowerPoint presentation, success message shown   │
  │      │                                      │ in chat                                                           │
  └──────┴──────────────────────────────────────┴───────────────────────────────────────────────────────────────────┘

  ---
  4. Vocabulary — With Native Language

  Settings: English, B1, Native Language: German

  ┌──────┬────────────────────────────────────────────────────┬─────────────────────────────────────────────────────┐
  │ Step │                       Action                       │                   Expected Result                   │
  ├──────┼────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ 4.1  │ Type: Give me 5 travel vocabulary words with       │ Slides generated                                    │
  │      │ examples                                           │                                                     │
  ├──────┼────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ 4.2  │ Check each vocabulary slide                        │ Subtitle field contains German translation of the   │
  │      │                                                    │ word                                                │
  └──────┴────────────────────────────────────────────────────┴─────────────────────────────────────────────────────┘

  ---
  5. Vocabulary — Requirements Validation

  Goal: Confirm the AI asks for missing info instead of guessing.

  ┌──────┬───────────────────────────────────────────────────────────┬──────────────────────────────────────────────┐
  │ Step │                           Input                           │             Expected AI Response             │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 5.1  │ Give me some vocabulary words                             │ AI asks: how many words? what topic?         │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 5.2  │ Give me food vocabulary                                   │ AI asks: how many words? include examples?   │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 5.3  │ Give me 5 words                                           │ AI asks: what topic? include examples?       │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 5.4  │ Give me 5 food words                                      │ AI asks: should examples be included?        │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 5.5  │ Answer the clarification, e.g. yes, include 1 example per │ AI generates with the now-complete           │
  │      │  word                                                     │ requirements                                 │
  └──────┴───────────────────────────────────────────────────────────┴──────────────────────────────────────────────┘

  ---
  6. Grammar — Happy Path

  Settings: English, B1

  ┌──────┬───────────────────────────────────────────────────────┬──────────────────────────────────────────────────┐
  │ Step │                        Action                         │                 Expected Result                  │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 6.1  │ Type: Create 3 grammar slides about present perfect   │ Slides generated                                 │
  │      │ with examples                                         │                                                  │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 6.2  │ Check title slide                                     │ Shows "Present Perfect" title and subtitle       │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 6.3  │ Check content slides                                  │ Short lines, symbols (📌 🔹 ✓ ✗) used for visual │
  │      │                                                       │  structure                                       │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 6.4  │ Check example sections                                │ Example sentences visible in example field       │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 6.5  │ Navigate through all slides and insert                │ All slides appear in PowerPoint                  │
  └──────┴───────────────────────────────────────────────────────┴──────────────────────────────────────────────────┘

  ---
  7. Grammar — Requirements Validation

  ┌──────┬───────────────────────────────────────┬───────────────────────────────────────────────────────┐
  │ Step │                 Input                 │                 Expected AI Response                  │
  ├──────┼───────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 7.1  │ Make grammar slides about past simple │ Asks: how many slides? include examples?              │
  ├──────┼───────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 7.2  │ Create 2 grammar slides               │ Asks: what grammar topic?                             │
  ├──────┼───────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 7.3  │ Give me grammar with examples         │ Asks: what topic? how many slides? how many examples? │
  └──────┴───────────────────────────────────────┴───────────────────────────────────────────────────────┘

  ---
  8. Quiz — Happy Path

  Settings: English, B1

  Run this after completing a vocabulary session so there's conversation context.

  ┌──────┬───────────────────────────────────────────────────────────┬──────────────────────────────────────────────┐
  │ Step │                          Action                           │               Expected Result                │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.1  │ First generate: Give me 5 food vocabulary words with      │ Vocabulary in context                        │
  │      │ examples → insert                                         │                                              │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.2  │ Type: Create a 5-question multiple-choice quiz on         │ Quiz slides generated                        │
  │      │ vocabulary                                                │                                              │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.3  │ Check quiz slides                                         │ Each slide has a question + A/B/C/D options  │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.4  │ Navigate and insert                                       │ Quiz slides added to presentation            │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.5  │ Start fresh (New Chat), then request: Create a 3-question │ AI asks: what topic (vocabulary or grammar)? │
  │      │  true-false quiz                                          │  what focus?                                 │
  ├──────┼───────────────────────────────────────────────────────────┼──────────────────────────────────────────────┤
  │ 8.6  │ Try: 5-question fill-in-the-blank quiz on vocabulary food │ AI generates fill-in-the-blank format with   │
  │      │  words                                                    │ ___ blank                                    │
  └──────┴───────────────────────────────────────────────────────────┴──────────────────────────────────────────────┘

  ---
  9. Quiz — Requirements Validation

  ┌──────┬──────────────────────────────────────────────┬───────────────────────────────────────────────────────────┐
  │ Step │                    Input                     │                   Expected AI Response                    │
  ├──────┼──────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
  │ 9.1  │ Make a quiz                                  │ Asks: topic, quiz type, number of questions               │
  ├──────┼──────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
  │ 9.2  │ 5-question quiz about food                   │ Asks: what type (multiple-choice, true-false,             │
  │      │                                              │ fill-in-the-blank)?                                       │
  ├──────┼──────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
  │ 9.3  │ Multiple-choice quiz about food              │ Asks: how many questions?                                 │
  ├──────┼──────────────────────────────────────────────┼───────────────────────────────────────────────────────────┤
  │ 9.4  │ Context has both vocab + grammar, request is │ Asks: focus on vocabulary or grammar?                     │
  │      │  vague                                       │                                                           │
  └──────┴──────────────────────────────────────────────┴───────────────────────────────────────────────────────────┘

  ---
  10. Homework — Happy Path

  Settings: English, B1 — run after vocabulary session.

  ┌──────┬────────────────────────────────────────────────────────────────────┬─────────────────────────────────────┐
  │ Step │                               Action                               │           Expected Result           │
  ├──────┼────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 10.1 │ After vocabulary slides in context, type: Create 5                 │ Homework slides generated           │
  │      │ fill-in-the-blank homework tasks on vocabulary                     │                                     │
  ├──────┼────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 10.2 │ Check slides                                                       │ Instruction shown, numbered items   │
  │      │                                                                    │ with ___ blanks                     │
  ├──────┼────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 10.3 │ Try: Create 3 sentence transformation homework tasks on grammar    │ Works if grammar is in context      │
  ├──────┼────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
  │ 10.4 │ Try: Create 5 matching homework tasks on vocabulary                │ Matching-style tasks generated      │
  └──────┴────────────────────────────────────────────────────────────────────┴─────────────────────────────────────┘

  ---
  11. Homework — Requirements Validation

  ┌──────┬────────────────────────────────────────────────┬───────────────────────────────────────────────────────┐
  │ Step │                     Input                      │                 Expected AI Response                  │
  ├──────┼────────────────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 11.1 │ Create homework                                │ Asks: type, number of tasks, focus (vocab or grammar) │
  ├──────┼────────────────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 11.2 │ Create fill-in-the-blank homework              │ Asks: how many tasks? vocab or grammar focus?         │
  ├──────┼────────────────────────────────────────────────┼───────────────────────────────────────────────────────┤
  │ 11.3 │ Both vocab + grammar in context, vague request │ Asks: focus on vocabulary or grammar?                 │
  └──────┴────────────────────────────────────────────────┴───────────────────────────────────────────────────────┘

  ---
  12. Preview Navigation

  Goal: Test all navigation controls.

  ┌──────┬───────────────────────────────────┬───────────────────────────────────────────────────────────────┐
  │ Step │              Action               │                        Expected Result                        │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.1 │ On slide 1, click Back button     │ Button is disabled (greyed out)                               │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.2 │ Press Backspace key               │ Same — no navigation on first slide                           │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.3 │ Click Next or press Enter         │ Moves to slide 2, Back button becomes enabled                 │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.4 │ Press Backspace                   │ Goes back to slide 1                                          │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.5 │ Click Remove (R) on a slide       │ Slide is removed from list, counter updates                   │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.6 │ Remove all slides                 │ Preview disappears, no slides to insert                       │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.7 │ Press R keyboard shortcut         │ Same as clicking Remove                                       │
  ├──────┼───────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 12.8 │ Click the × in the preview header │ Preview cancelled, "Preview cancelled." message shown in chat │
  └──────┴───────────────────────────────────┴───────────────────────────────────────────────────────────────┘

  ---
  13. Edit a Slide

  Goal: Test per-slide editing with natural language.

  ┌──────┬──────────────────────────────────────────────────────┬───────────────────────────────────────────────────┐
  │ Step │                        Action                        │                  Expected Result                  │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.1 │ During preview of vocabulary, navigate to a word     │ Slide shown                                       │
  │      │ slide                                                │                                                   │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.2 │ Click Edit (E) or press E                            │ "Editing slide N" badge appears above input box   │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.3 │ Type: Make the definition shorter                    │ Progress spinner, then preview updates with new   │
  │      │                                                      │ definition                                        │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.4 │ Type: Change the example sentence to use a           │ Example sentence updated                          │
  │      │ restaurant context                                   │                                                   │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.5 │ Type: Replace this word with "recipe"                │ Word, definition, and example all updated         │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.6 │ Click the × on the edit badge                        │ Exit edit mode, badge disappears                  │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.7 │ Edit a grammar slide with: Simplify the explanation  │ Grammar content updated                           │
  ├──────┼──────────────────────────────────────────────────────┼───────────────────────────────────────────────────┤
  │ 13.8 │ Edit a quiz slide with: Change option B to be more   │ Quiz options updated                              │
  │      │ plausible                                            │                                                   │
  └──────┴──────────────────────────────────────────────────────┴───────────────────────────────────────────────────┘

  ---
  14. Cancel During Generation

  ┌──────┬───────────────────────────────────────────────────────┬──────────────────────────────────────────────────┐
  │ Step │                        Action                         │                 Expected Result                  │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 14.1 │ Send any request, immediately click × on the progress │ "Generation cancelled." message shown, no slides │
  │      │  spinner                                              │  appear                                          │
  ├──────┼───────────────────────────────────────────────────────┼──────────────────────────────────────────────────┤
  │ 14.2 │ Send a new request after cancelling                   │ Works normally, previous cancel is ignored       │
  └──────┴───────────────────────────────────────────────────────┴──────────────────────────────────────────────────┘

  ---
  15. Multi-Turn Conversation (Full Lesson Flow)

  Goal: Test that context carries across all content types in a single session.

  ┌──────┬─────────────────────────────────────────────────┬────────────────────────────────────────────────────────┐
  │ Step │                     Action                      │                        Expected                        │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.1 │ Give me 6 travel vocabulary words with 1        │ Vocabulary generated                                   │
  │      │ example each                                    │                                                        │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.2 │ Insert slides                                   │ Vocabulary in conversation context                     │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.3 │ Create 3 grammar slides about modal verbs with  │ Grammar generated, examples relate to travel context   │
  │      │ examples                                        │ if possible                                            │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.4 │ Insert slides                                   │ Grammar in context                                     │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.5 │ Create a 5-question multiple-choice quiz on     │ Quiz uses the travel words from step 15.1              │
  │      │ vocabulary                                      │                                                        │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.6 │ Insert slides                                   │ Quiz in context                                        │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.7 │ Create 5 fill-in-the-blank homework tasks on    │ Homework uses travel vocabulary                        │
  │      │ vocabulary                                      │                                                        │
  ├──────┼─────────────────────────────────────────────────┼────────────────────────────────────────────────────────┤
  │ 15.8 │ Insert all                                      │ Complete lesson deck in PowerPoint                     │
  └──────┴─────────────────────────────────────────────────┴────────────────────────────────────────────────────────┘

  ---
  16. New Chat / Reset

  ┌──────┬──────────────────────────────────────────────┬────────────────────────────────────────────────────────────┐
  │ Step │                    Action                    │                      Expected Result                       │
  ├──────┼──────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 16.1 │ Click the New Chat button (+ icon in header) │ Chat clears, welcome screen returns, conversation ID reset │
  ├──────┼──────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 16.2 │ Try requesting a quiz immediately            │ AI has no vocabulary context, asks for clarification       │
  ├──────┼──────────────────────────────────────────────┼────────────────────────────────────────────────────────────┤
  │ 16.3 │ Settings should be preserved after new chat  │ Badge still shows same language/level                      │
  └──────┴──────────────────────────────────────────────┴────────────────────────────────────────────────────────────┘

  ---
  17. Multilingual Teacher Input

  Goal: Teacher can write in their own language, content is generated in target language.

  ┌──────┬───────────────────────────────────────────────────────────────────────┬──────────────────────────────────┐
  │ Step │                                Action                                 │             Expected             │
  ├──────┼───────────────────────────────────────────────────────────────────────┼──────────────────────────────────┤
  │ 17.1 │ Set target language: English, B1. Type request in German: Gib mir 5   │ Content generated in English     │
  │      │ Vokabeln zum Thema Essen mit Beispielen                               │                                  │
  ├──────┼───────────────────────────────────────────────────────────────────────┼──────────────────────────────────┤
  │ 17.2 │ If requirements are missing, AI responds in German                    │ Requirements-not-met message is  │
  │      │                                                                       │ in German                        │
  ├──────┼───────────────────────────────────────────────────────────────────────┼──────────────────────────────────┤
  │ 17.3 │ Set target language: German. Type request in English: Give me 5 food  │ Content generated in German      │
  │      │ words with examples                                                   │                                  │
  └──────┴───────────────────────────────────────────────────────────────────────┴──────────────────────────────────┘

  ---
  18. Edge Cases

  ┌──────┬──────────────────────────────────────────┬───────────────────────────────────────────────────────────────┐
  │ Step │                 Scenario                 │                           Expected                            │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.1 │ Submit empty message                     │ Nothing sent, no error                                        │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.2 │ Send a second request while preview is   │ Previous preview dismissed with "N slides not inserted", new  │
  │      │ still open                               │ request starts                                                │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.3 │ Request with age group set to "8-10"     │ Simpler vocabulary and definitions                            │
  │      │ (children)                               │                                                               │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.4 │ Request with age group set to "18+"      │ More sophisticated content                                    │
  │      │ (adults)                                 │                                                               │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.5 │ Backend is offline / not reachable       │ Error message: "Not connected to server. Make sure the        │
  │      │                                          │ backend is running."                                          │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.6 │ WebSocket disconnects mid-generation     │ Auto-reconnect up to 3 times                                  │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.7 │ Request 10+ vocabulary words             │ All words appear as separate slides + 1 title slide           │
  ├──────┼──────────────────────────────────────────┼───────────────────────────────────────────────────────────────┤
  │ 18.8 │ Open add-in in PowerPoint Online         │ slides.add() may not work — known limitation, document it     │
  │      │ (browser)                                │                                                               │
  └──────┴──────────────────────────────────────────┴───────────────────────────────────────────────────────────────┘

  ---
  Quick Reference: Keyboard Shortcuts

  ┌─────────────────┬───────────────────────────┐
  │       Key       │          Action           │
  ├─────────────────┼───────────────────────────┤
  │ Enter           │ Send message / Next slide │
  ├─────────────────┼───────────────────────────┤
  │ Backspace       │ Previous slide            │
  ├─────────────────┼───────────────────────────┤
  │ → (Right Arrow) │ Next slide                │
  ├─────────────────┼───────────────────────────┤
  │ E               │ Edit current slide        │
  ├─────────────────┼───────────────────────────┤
  │ R               │ Remove current slide      │
  ├─────────────────┼───────────────────────────┤
  │ Q               │ Cancel generation         │
  ├─────────────────┼───────────────────────────┤
  │ A               │ Insert all slides         │
  └─────────────────┴───────────────────────────┘

  ---
  What to Log When You Find Issues

  For each bug found, note:
  1. Settings at the time (language, level, native lang, age group)
  2. Exact prompt typed
  3. Content type selected (vocabulary/grammar/quiz/homework)
  4. What happened vs what was expected
  5. Was there conversation context (prior slides in the same session)?
  
  
  Click on the X does not closes / created new coversation it should be like with + 
  also delete -> all slides removed 
  alerting and monitoring 
	-> somehting bad happen I want nice message, good log, automaticall alert (email or something) 
  saving what users are requesting 
	-> mayne in the file for the first version 

BE and FE clean dead code 	
	