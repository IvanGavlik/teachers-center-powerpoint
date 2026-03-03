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

  ---
  TC-02 — Topic given but no slide count
  I need vocabulary slides about animals

  Expected: GPT asks for the number of slides

  ---
  TC-03 — Slide count given but no topic
  Give me 5 slides

  Expected: GPT asks what topic/content the slides should be about

  ---
  TC-04 — Completely vague
  Help me make a lesson

  Expected: GPT asks what type of content (explanation, task, quiz?), what topic, how many slides

  ---
  TC-05 — Wrong content implied (no explicit type)
  I want something about Present Simple

  Expected: GPT asks — do you want explanation, examples, tasks, or a quiz on Present Simple?

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

  ---
  TC-07 — Grammar explanation
  I need 2 slides explaining Present Simple affirmative and negative forms

  Expected:
  - 2 slides, one for affirmative, one for negative
  - Content uses transformation arrows (→), pattern formatting
  - Rule + example format, no mixing of too many concepts per slide

  ---
  TC-08 — Task/exercise type (not explanation)
  I need 3 task slides on Present Simple, use food vocabulary in the tasks

  Expected:
  - 3 slides with tasks/exercises
  - NOT explanations or definitions — actual student tasks
  - Food-related sentences

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

  ---
  TC-13 — Native language enabled

  Settings: Language = English, Level = A1, Native Language = German
  I need 3 vocabulary slides about colours

  Expected:
  - Slides generated
  - Each slide has a translation field with German translations

  ---
  TC-14 — Native language = No (explicit check)

  Settings: Native Language = No
  I need 2 vocabulary slides about weather

  Expected: No translation field anywhere in the response

  ---
  PATH 4 — Conversation Flow (multi-turn)

  These test that GPT remembers context from earlier in the chat.

  ---
  TC-15 — GPT asks, teacher answers, content generates

  Turn 1: I need vocabulary slides about sports
  → Expected: GPT asks how many slides

  Turn 2: 3 slides please
  → Expected: 3 slides about sports vocabulary generated

  ---
  TC-16 — Start new chat resets context

  After TC-15, click "New Chat" then:
  Give me 3 slides please

  Expected: GPT asks for topic again — it does NOT remember the previous "sports" context

  ---
  PATH 5 — Slide Editing

  These require slides to already be generated and visible in preview.

  ---
  TC-17 — Edit content of one slide

  First generate slides (e.g. TC-06).
  Then click edit on slide 1 and type:
  Make the content shorter, maximum 4 lines

  Expected: That single slide returns with condensed content, other slides unchanged

  ---
  TC-18 — Change a slide's topic

  Generate slides, edit slide 2:
  Change the topic of this slide to drinks instead of food

  Expected: Slide 2 regenerated with drinks vocabulary, same format

  ---
  TC-19 — Add examples to a slide

  Generate a grammar slide, edit it:
  Add 2 example sentences showing the rule in action

  Expected: Same slide returned with 2 new example sentences added

  ---
  TC-20 — Make slide simpler (level adjustment)

  Edit a slide:
  Simplify this for A1 students, use very basic vocabulary

  Expected: Simpler language, shorter words, no complex structures

  ---
  PATH 6 — Edge Cases & Stress

  ---
  TC-21 — Request irrelevant to language teaching
  Can you write me a business email?

  Expected: GPT should either refuse politely or ask to clarify in the context of language teaching (the system prompt
  constrains it to language teacher role)

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

  ┌────────────────────────────────────────────┬─────────────────────────────────────────────────┐
  │                   Signal                   │                  What it means                  │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ requirements-not-met field in response     │ Prompt validation working correctly             │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ Slides with 5–7 lines max                  │ Cognitive load rule respected                   │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ Symbols used (✓ ✗ →)                       │ Formatting style applied                        │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ No long paragraphs                         │ Scannability rule respected                     │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ One concept per slide                      │ "Break complex into small units" rule respected │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ translation field present/absent           │ Native language rule applied                    │
  ├────────────────────────────────────────────┼─────────────────────────────────────────────────┤
  │ Response language matches request language │ Language-awareness working                      │
  └────────────────────────────────────────────┴─────────────────────────────────────────────────┘