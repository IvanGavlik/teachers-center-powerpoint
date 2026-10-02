/* global Office, PowerPoint */

/**
 * Teacher Assistant Taskpane - Full chat/preview UI with direct PowerPoint API access.
 * Ported from dialog.js to run in the taskpane sidebar (~350px).
 * No parent-child messaging needed — calls PowerPoint.run() directly.
 */

// Import CSS for webpack bundling
import './taskpane.css';
import PptxGenJS from 'pptxgenjs';
import QRCode from 'qrcode';
import { detectSlideSize, readDocumentZip, toSlideFormat } from './slides/slideSize';
import { buildDeckBase64, insertDeckBase64 } from './slides/renderDeck';
import { renderSlideCard, cardLabel } from './slides/previewCards';

// ============================================
// WEBSOCKET CONFIGURATION
// ============================================

const WS_URL = process.env.WS_URL;
const API_URL = process.env.API_URL;
const GAME_BASE_URL = `${API_URL}/game`;
const USER_ID = 'user-123';  // TODO: implement proper user management
const CHANNEL_NAME = 'powerpoint-taskpane';
const MAX_RECONNECT_ATTEMPTS = 3;

const COMMANDS = [
    { command: '/multiple-choice', description: 'Multiple choice quiz', icon: 'sports_esports', mode: 'multiple-choice' },
    { command: '/sentence-ordering', description: 'Sentence ordering game', icon: 'reorder', mode: 'sentence-ordering' },
    { command: '/guess-the-word', description: 'Guess the word', icon: 'psychology', mode: 'guess-the-word' },
    { command: '/true-false', description: 'True or false game', icon: 'fact_check', mode: 'true-false' }
];

const INTERACTIVITY_KEYWORDS = [
    { keyword: 'multiple choice', mode: 'multiple-choice' },
    { keyword: 'mcq', mode: 'multiple-choice' },
    { keyword: 'quiz game', mode: 'multiple-choice' },
    { keyword: 'interactive quiz', mode: 'multiple-choice' },
    { keyword: 'make a quiz', mode: 'multiple-choice' },
    { keyword: 'create a quiz', mode: 'multiple-choice' },
    { keyword: 'quiz activity', mode: 'multiple-choice' },
    { keyword: 'class quiz', mode: 'multiple-choice' },
    { keyword: 'student quiz', mode: 'multiple-choice' },
    { keyword: 'practice quiz', mode: 'multiple-choice' },
    { keyword: 'sentence order', mode: 'sentence-ordering' },
    { keyword: 'word order', mode: 'sentence-ordering' },
    { keyword: 'unscramble', mode: 'sentence-ordering' },
    { keyword: 'arrange the words', mode: 'sentence-ordering' },
    { keyword: 'sentence scramble', mode: 'sentence-ordering' },
    { keyword: 'order the words', mode: 'sentence-ordering' },
    { keyword: 'guess the word', mode: 'guess-the-word' },
    { keyword: 'word from definition', mode: 'guess-the-word' },
    { keyword: 'guess the vocabulary', mode: 'guess-the-word' },
    { keyword: 'definition game', mode: 'guess-the-word' },
    { keyword: 'true or false', mode: 'true-false' },
    { keyword: 'true/false', mode: 'true-false' },
    { keyword: 'true false quiz', mode: 'true-false' },
    { keyword: 'fact check', mode: 'true-false' }
];

const INTERACTIVITY_MODE_LABELS = {
    'multiple-choice': {
        label: 'Multiple Choice Quiz',
        icon: 'sports_esports',
        placeholder: 'Describe your quiz topic...',
        successMessage: 'Quiz slide added! Students can scan the QR code to play.',
        countNoun: 'questions',
        joinStep3: 'Answer the questions'
    },
    'sentence-ordering': {
        label: 'Sentence Ordering',
        icon: 'reorder',
        placeholder: 'Describe the sentences/topic to practice...',
        successMessage: 'Sentence ordering slide added! Students can scan the QR code to play.',
        countNoun: 'sentences',
        joinStep3: 'Arrange the words in order'
    },
    'guess-the-word': {
        label: 'Guess the Word',
        icon: 'psychology',
        placeholder: 'Describe the vocabulary/topic to practice...',
        successMessage: 'Guess the word slide added! Students can scan the QR code to play.',
        countNoun: 'words',
        joinStep3: 'Guess the word from its definition'
    },
    'true-false': {
        label: 'True or False',
        icon: 'fact_check',
        placeholder: 'Describe the topic for true/false statements...',
        successMessage: 'True or false slide added! Students can scan the QR code to play.',
        countNoun: 'statements',
        joinStep3: 'Decide if each statement is true or false'
    }
};

// ============================================
// SLIDE THEME — single source of truth for colors used in
// both the preview UI (injected as CSS vars) and slide insertion
// ============================================

const SLIDE_THEME = {
    colors: {
        title:        '#d13438',   // --accent-primary  (slide titles, accent1 from theme)
        accent2:      '#ed7d31',   // --accent-2
        accent3:      '#a5a5a5',   // --accent-3
        accent4:      '#ffc000',   // --accent-4
        accent5:      '#5b9bd5',   // --accent-5
        accent6:      '#70ad47',   // --accent-6
        accentHover:  '#b7472a',   // --accent-hover
        accentLight:  '#fdf3f3',   // --accent-light
        subtitle:     '#605e5c',   // --text-secondary  (subtitles, examples)
        content:      '#323130',   // --text-primary    (body text)
        bgSlide:      '#ffffff',   // --bg-slide        (slide background, background1 from theme)
        bgSlideAlt:   '#f3f2f1',   // --bg-slide-alt    (secondary slide bg, background2 from theme)
        bgPrimary:    '#faf9f8',   // --bg-primary
        bgSecondary:  '#ffffff',   // --bg-secondary
        bgTertiary:   '#f3f2f1',   // --bg-tertiary
        borderLight:  '#edebe9',   // --border-light
        borderMedium: '#d2d0ce',   // --border-medium
        successBg:    '#dff6dd',   // --success-bg
        successText:  '#107c10',   // --success-text
        errorBg:      '#fde7e9',   // --error-bg
        errorText:    '#a80000',   // --error-text
    },
    fonts: {
        heading: 'Calibri Light',  // --font-heading (title shapes, majorFont from theme)
        body:    'Calibri'         // --font-body    (content/subtitle/example, minorFont from theme)
    }
};

// ============================================
// STATE MANAGEMENT
// ============================================

const state = {
    // Chat state
    isProcessing: false,
    pendingRequest: null,

    // Preview state
    deck: null,              // { title, slides: [{ layout, … }] } from the backend — source of truth for layout slides
    slides: [],              // preview cards: derived from `deck` ({ kind: 'title' | 'layout' }), or interactivity cards ({ title, content, type })
    currentSlideIndex: 0,
    isInPreviewMode: false,

    // Deck slide size in inches ({ w, h, source }) — layout hint for the backend, real size for insert
    slideSize: null,

    // Edit mode state
    originalRequest: null,
    editingSlideIndex: null,
    isEditMode: false,

    // Progress state
    progressElement: null,
    cancelled: false,
    staleResponses: 0,


    // Settings
    settingsConfirmed: false,
    settings: {
        language: 'English',
        level: 'B1',
        ageGroup: ''
    },

    // Current file path (for per-file settings)
    currentFilePath: null,

    // Platform
    isWeb: false,

    // Interactivity mode
    interactivityMode: null,           // null | 'multiple-choice' | 'sentence-ordering'
    pendingInteractivityRequest: null, // message stored while waiting for confirmation
    pendingInteractivityMode: null,    // mode detected for the pending confirmation
    pendingActivity: null,             // full activity payload while in interactivity preview

    // Attached book (file_search grounding) — one at a time, ephemeral to this session
    attachedBook: null,                // null | { vectorStoreId, filename, status, sent }
                                        // status: 'uploading' | 'processing' | 'ready' | 'error'
    bookChangeConfirmEl: null,         // open "remove the document?" bubble, if any
    bookPollTimer: null,
    bookPollAttempts: 0,

    // WebSocket state
    ws: null,
    wsState: 'disconnected',
    reconnectAttempts: 0,
    conversationId: null,

    // Preview element reference
    previewElement: null,

    // UI elements (cached after init)
    elements: {}
};

// Selected star value for the feedback modal (module-level so setupEventListeners can access it)
let selectedStarRating = 0;

// ============================================
// INITIALIZATION
// ============================================

// ============================================
// OFFICE INITIALIZATION
// ============================================

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        state.isWeb = Office.context.platform === Office.PlatformType.OfficeOnline;
        document.getElementById('sideload-msg').style.display = 'none';
        document.getElementById('app-body').style.display = 'flex';
        initializeTaskpane();
    }
});

function applyCSSVariables() {
    const root = document.documentElement.style;
    root.setProperty('--accent-primary',  SLIDE_THEME.colors.title);
    root.setProperty('--accent-2',        SLIDE_THEME.colors.accent2);
    root.setProperty('--accent-3',        SLIDE_THEME.colors.accent3);
    root.setProperty('--accent-4',        SLIDE_THEME.colors.accent4);
    root.setProperty('--accent-5',        SLIDE_THEME.colors.accent5);
    root.setProperty('--accent-6',        SLIDE_THEME.colors.accent6);
    root.setProperty('--accent-hover',    SLIDE_THEME.colors.accentHover);
    root.setProperty('--accent-light',    SLIDE_THEME.colors.accentLight);
    root.setProperty('--text-primary',    SLIDE_THEME.colors.content);
    root.setProperty('--text-secondary',  SLIDE_THEME.colors.subtitle);
    root.setProperty('--bg-slide',        SLIDE_THEME.colors.bgSlide);
    root.setProperty('--bg-slide-alt',    SLIDE_THEME.colors.bgSlideAlt);
    root.setProperty('--bg-primary',      SLIDE_THEME.colors.bgPrimary);
    root.setProperty('--bg-secondary',    SLIDE_THEME.colors.bgSecondary);
    root.setProperty('--bg-tertiary',     SLIDE_THEME.colors.bgTertiary);
    root.setProperty('--border-light',    SLIDE_THEME.colors.borderLight);
    root.setProperty('--border-medium',   SLIDE_THEME.colors.borderMedium);
    root.setProperty('--success-bg',      SLIDE_THEME.colors.successBg);
    root.setProperty('--success-text',    SLIDE_THEME.colors.successText);
    root.setProperty('--error-bg',        SLIDE_THEME.colors.errorBg);
    root.setProperty('--error-text',      SLIDE_THEME.colors.errorText);
    root.setProperty('--font-heading',    SLIDE_THEME.fonts.heading);
    root.setProperty('--font-body',       SLIDE_THEME.fonts.body);
}

async function readPresentationTheme() {
    if (state.isWeb) {
        // getFileAsync(FileType.Compressed) is not supported in PowerPoint for the web.
        // Apply CSS variables from the current SLIDE_THEME defaults and return.
        console.log('[Theme] Web platform — skipping file read, using default theme.');
        applyCSSVariables();
        return;
    }
    console.log('[Theme] Reading presentation theme from file...');
    try {
        const theme = await readThemeFromFile();
        console.log('[Theme] Raw parsed theme:', theme);

        SLIDE_THEME.colors.title      = theme.accent1 || SLIDE_THEME.colors.title;
        SLIDE_THEME.colors.accent2    = theme.accent2 || SLIDE_THEME.colors.accent2;
        SLIDE_THEME.colors.accent3    = theme.accent3 || SLIDE_THEME.colors.accent3;
        SLIDE_THEME.colors.accent4    = theme.accent4 || SLIDE_THEME.colors.accent4;
        SLIDE_THEME.colors.accent5    = theme.accent5 || SLIDE_THEME.colors.accent5;
        SLIDE_THEME.colors.accent6    = theme.accent6 || SLIDE_THEME.colors.accent6;
        SLIDE_THEME.colors.content    = theme.dark1   || SLIDE_THEME.colors.content;
        SLIDE_THEME.colors.subtitle   = theme.dark2   || SLIDE_THEME.colors.subtitle;
        SLIDE_THEME.colors.bgSlide    = theme.light1  || SLIDE_THEME.colors.bgSlide;
        SLIDE_THEME.colors.bgSlideAlt = theme.light2  || SLIDE_THEME.colors.bgSlideAlt;

        SLIDE_THEME.fonts.heading = theme.majorFont || SLIDE_THEME.fonts.heading;
        SLIDE_THEME.fonts.body    = theme.minorFont || SLIDE_THEME.fonts.body;

        console.log('[Theme] Applied theme — colors:', {
            title: SLIDE_THEME.colors.title,
            content: SLIDE_THEME.colors.content,
            subtitle: SLIDE_THEME.colors.subtitle,
            bgSlide: SLIDE_THEME.colors.bgSlide,
        }, 'fonts:', SLIDE_THEME.fonts);

        applyCSSVariables();
        console.log('[Theme] CSS variables updated.');
    } catch (err) {
        console.warn('[Theme] Could not read presentation theme, using defaults:', err);
    }
}

// Reads the PPTX file via Office API, extracts ppt/theme/theme1.xml from the ZIP,
// and parses both the color scheme and font scheme from it. Both getThemeColorsAsync
// and getThemeFontsAsync are not implemented in current PowerPoint builds.
// The theme tints the taskpane/preview and gives the slide text measurer the deck's fonts.
async function readThemeFromFile() {
    const zip = await readDocumentZip();
    const themeEntry = zip.file('ppt/theme/theme1.xml');
    if (!themeEntry) {
        console.warn('[Theme] ppt/theme/theme1.xml not found in ZIP.');
        return {};
    }
    const xml = await themeEntry.async('string');
    console.log(`[Theme] theme1.xml extracted (${xml.length} chars).`);
    return parseThemeXml(xml);
}

// Extracts colors and fonts from OOXML theme XML.
// Colors: handles both <a:srgbClr val="RRGGBB"> and <a:sysClr lastClr="RRGGBB">.
// Fonts: reads <a:latin typeface="..."> inside majorFont and minorFont sections.
function parseThemeXml(xml) {
    // Two-step: first extract the element's inner content, then search within it.
    // This prevents the non-greedy [\s\S]*? from crossing element boundaries when
    // an element uses sysClr (no srgbClr) and the next element happens to have one.
    const extractColor = (tag) => {
        const elem = new RegExp(`<a:${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</a:${tag}>`, 'i').exec(xml);
        if (!elem) return null;
        const inner = elem[1];
        const srgb = /<a:srgbClr val="([0-9a-fA-F]{6})"/i.exec(inner);
        if (srgb) return '#' + srgb[1];
        const sys = /<a:sysClr[^>]*lastClr="([0-9a-fA-F]{6})"/i.exec(inner);
        if (sys) return '#' + sys[1];
        return null;
    };

    const extractFont = (section) => {
        const elem = new RegExp(`<a:${section}(?:\\s[^>]*)?>([\\s\\S]*?)</a:${section}>`, 'i').exec(xml);
        if (!elem) return null;
        const match = /<a:latin typeface="([^"]+)"/i.exec(elem[1]);
        return match ? match[1] : null;
    };

    const result = {
        dark1:     extractColor('dk1'),
        light1:    extractColor('lt1'),
        dark2:     extractColor('dk2'),
        light2:    extractColor('lt2'),
        accent1:   extractColor('accent1'),
        accent2:   extractColor('accent2'),
        accent3:   extractColor('accent3'),
        accent4:   extractColor('accent4'),
        accent5:   extractColor('accent5'),
        accent6:   extractColor('accent6'),
        majorFont: extractFont('majorFont'),
        minorFont: extractFont('minorFont'),
    };

    const missing = Object.entries(result).filter(([, v]) => v === null).map(([k]) => k);
    if (missing.length) console.warn('[Theme] Could not extract from XML:', missing.join(', '));

    return result;
}

function initializeTaskpane() {
    console.log('[Init] Taskpane initializing. Default theme colors:', SLIDE_THEME.colors.title, '| fonts:', SLIDE_THEME.fonts);
    applyCSSVariables();
    refreshSlideSize();

    // Cache DOM elements
    state.elements = {
        chatBody: document.getElementById('chatBody'),
        welcomeState: document.getElementById('welcomeState'),
        messageInput: document.getElementById('messageInput'),
        inputContainer: document.getElementById('inputContainer'),
        newChatBtn: document.getElementById('newChatBtn'),
        contextBadge: document.getElementById('contextBadge'),
        editBadge: document.getElementById('editBadge'),
        editBadgeSlideNum: document.getElementById('editBadgeSlideNum'),
        editBadgeCancelBtn: document.getElementById('editBadgeCancelBtn'),
        interactivityChip: document.getElementById('interactivityChip'),
        interactivityChipIcon: document.getElementById('interactivityChipIcon'),
        interactivityChipLabel: document.getElementById('interactivityChipLabel'),
        interactivityChipCancelBtn: document.getElementById('interactivityChipCancelBtn'),
        attachBookBtn: document.getElementById('attachBookBtn'),
        bookFileInput: document.getElementById('bookFileInput'),
        bookChip: document.getElementById('bookChip'),
        bookChipIcon: document.getElementById('bookChipIcon'),
        bookChipLabel: document.getElementById('bookChipLabel'),
        bookChipStatus: document.getElementById('bookChipStatus'),
        bookChipCancelBtn: document.getElementById('bookChipCancelBtn'),
        commandsAutocomplete: document.getElementById('commandsAutocomplete'),
        settingsModal: document.getElementById('settingsModal'),
        closeModalBtn: document.getElementById('closeModalBtn'),
        cancelSettingsBtn: document.getElementById('cancelSettingsBtn'),
        saveSettingsBtn: document.getElementById('saveSettingsBtn'),
        settingsLanguage: document.getElementById('settingsLanguage'),
        settingsLevel: document.getElementById('settingsLevel'),
        settingsAgeGroup: document.getElementById('settingsAgeGroup'),
        whatsNewModal: document.getElementById('whatsNewModal'),
        closeWhatsNewModalBtn: document.getElementById('closeWhatsNewModalBtn'),
        whatsNewCloseBtn: document.getElementById('whatsNewCloseBtn'),
        whatsNewDontShowAgain: document.getElementById('whatsNewDontShowAgain'),
        feedbackBtn: document.getElementById('feedbackBtn'),
        feedbackModal: document.getElementById('feedbackModal'),
        closeFeedbackModalBtn: document.getElementById('closeFeedbackModalBtn'),
        cancelFeedbackBtn: document.getElementById('cancelFeedbackBtn'),
        submitFeedbackBtn: document.getElementById('submitFeedbackBtn'),
        feedbackComment: document.getElementById('feedbackComment'),
        feedbackErrorContext: document.getElementById('feedbackErrorContext'),
        starRating: document.getElementById('starRating'),
        npsWidget: document.getElementById('npsWidget'),
        npsStarRating: document.getElementById('npsStarRating')
    };

    loadSettings();
    setupEventListeners();

    // Connect to WebSocket backend
    setTimeout(() => {
        connectWebSocket();
    }, 500);

    // Show NPS if threshold already met from previous sessions
    setTimeout(() => checkAndShowNPS(), 800);

    console.log('[Init] Taskpane ready.');
}

function setupEventListeners() {
    const { messageInput, newChatBtn } = state.elements;

    // Block input focus until settings are confirmed
    messageInput.addEventListener('focus', () => {
        if (!state.settingsConfirmed) {
            messageInput.blur();
            openSettingsModal();
        }
    });

    // Auto-resize textarea + command autocomplete
    messageInput.addEventListener('input', () => {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.min(messageInput.scrollHeight, 100) + 'px';

        const val = messageInput.value;
        if (val.startsWith('/')) {
            showCommandsAutocomplete(val);
        } else {
            hideCommandsAutocomplete();
        }
    });

    messageInput.addEventListener('keydown', handleCommandAutocompleteKeydown);

    messageInput.addEventListener('blur', () => {
        setTimeout(hideCommandsAutocomplete, 150);
    });

    // Header buttons
    newChatBtn.addEventListener('click', handleNewChat);

    // Edit badge cancel
    const { editBadgeCancelBtn } = state.elements;
    if (editBadgeCancelBtn) {
        editBadgeCancelBtn.addEventListener('click', exitEditMode);
    }

    // Interactivity chip cancel
    const { interactivityChipCancelBtn } = state.elements;
    if (interactivityChipCancelBtn) {
        interactivityChipCancelBtn.addEventListener('click', exitInteractivityMode);
    }

    // Attach-a-book (file_search grounding)
    const { attachBookBtn, bookFileInput, bookChipCancelBtn } = state.elements;
    if (attachBookBtn && bookFileInput) {
        attachBookBtn.addEventListener('click', () => bookFileInput.click());
        bookFileInput.addEventListener('change', () => {
            const file = bookFileInput.files && bookFileInput.files[0];
            if (file) handleBookFileSelected(file);
        });
    }
    if (bookChipCancelBtn) {
        bookChipCancelBtn.addEventListener('click', handleBookChipCancel);
    }

    // Context badge opens settings
    const { contextBadge } = state.elements;
    if (contextBadge) {
        contextBadge.addEventListener('click', openSettingsModal);
    }

    // Settings modal
    const { closeModalBtn, cancelSettingsBtn, saveSettingsBtn, settingsModal } = state.elements;
    closeModalBtn.addEventListener('click', closeSettingsModal);
    cancelSettingsBtn.addEventListener('click', closeSettingsModal);
    saveSettingsBtn.addEventListener('click', saveSettings);
    settingsModal.addEventListener('click', (e) => {
        if (e.target === settingsModal) closeSettingsModal();
    });

    // What's New modal
    const { closeWhatsNewModalBtn, whatsNewCloseBtn, whatsNewModal } = state.elements;
    closeWhatsNewModalBtn.addEventListener('click', closeWhatsNewModal);
    whatsNewCloseBtn.addEventListener('click', closeWhatsNewModal);
    whatsNewModal.addEventListener('click', (e) => {
        if (e.target === whatsNewModal) closeWhatsNewModal();
    });

    // Global keyboard shortcuts
    document.addEventListener('keydown', handleGlobalKeydown);

    // Feedback modal
    const { feedbackBtn, feedbackModal, closeFeedbackModalBtn, cancelFeedbackBtn, submitFeedbackBtn, starRating } = state.elements;
    feedbackBtn.addEventListener('click', () => openFeedbackModal());
    closeFeedbackModalBtn.addEventListener('click', closeFeedbackModal);
    cancelFeedbackBtn.addEventListener('click', closeFeedbackModal);
    submitFeedbackBtn.addEventListener('click', submitFeedbackModal);
    feedbackModal.addEventListener('click', (e) => {
        if (e.target === feedbackModal) closeFeedbackModal();
    });

    starRating.querySelectorAll('.star').forEach(star => {
        star.addEventListener('click', () => {
            selectedStarRating = parseInt(star.dataset.val);
            starRating.querySelectorAll('.star').forEach((s, i) => {
                s.classList.toggle('active', i < selectedStarRating);
            });
            state.elements.submitFeedbackBtn.disabled = false;
        });
    });
}

// ============================================
// MESSAGE HANDLING
// ============================================

// Slide size of the open deck — cached; refreshed at start-up, on New Chat and before every insert.
async function refreshSlideSize() {
    try {
        state.slideSize = await detectSlideSize(state.isWeb);
        console.log(`[SlideSize] ${state.slideSize.w.toFixed(2)}×${state.slideSize.h.toFixed(2)}in via ${state.slideSize.source}`);
    } catch (err) {
        console.warn('[SlideSize] detection failed, backend gets its 16:9 default:', err);
    }
    return state.slideSize;
}

// Dev builds only: /dev-deck [typical|max|overflow] — sample decks through the real preview + insert path.
async function runDevDeckCommand(content) {
    const [, variant = 'typical'] = content.split(/\s+/);
    state.elements.welcomeState.classList.add('hidden');
    addUserMessage(content);
    // Constant condition: webpack drops this block (and the fixtures chunk) from production builds.
    if (process.env.NODE_ENV !== 'production') {
        const { FIXTURES } = await import(/* webpackChunkName: "dev-fixtures" */ './slides/devFixtures');
        const fixture = FIXTURES[variant];
        if (!fixture) {
            showError(`Unknown sample deck "${variant}". Use typical, max or overflow.`);
            return;
        }
        await readPresentationTheme();
        showDeckPreview(JSON.parse(JSON.stringify(fixture)));
    }
}

function handleSend() {
    if (!state.settingsConfirmed) {
        openSettingsModal();
        return;
    }

    const { messageInput } = state.elements;
    const content = messageInput.value.trim();
    if (!content || state.isProcessing) return;
    if (isBookPending()) {
        flagBookPending();
        return;
    }

    // Sample decks for testing the renderer — dev builds only (branch removed in production builds)
    if (process.env.NODE_ENV !== 'production') {
        if (content.toLowerCase().startsWith('/dev-deck')) {
            messageInput.value = '';
            runDevDeckCommand(content);
            return;
        }
    }

    // Interactivity command detection — Option 1: explicit /command
    const matchedCommand = COMMANDS.find(c => content.toLowerCase().startsWith(c.command));
    if (matchedCommand) {
        messageInput.value = '';
        messageInput.style.height = 'auto';
        enterInteractivityMode(matchedCommand.mode);
        return;
    }

    // Interactivity keyword detection — Option 2: natural language
    const detectedMode = !state.interactivityMode ? detectInteractivityMode(content) : null;
    if (detectedMode) {
        messageInput.value = '';
        messageInput.style.height = 'auto';
        state.elements.welcomeState.classList.add('hidden');
        addUserMessage(content);
        showInteractivityConfirmation(content, detectedMode);
        return;
    }

    // Edit mode
    if (state.isEditMode && state.editingSlideIndex !== null) {
        handleEditSend(content);
        return;
    }

    // Dismiss existing preview
    if (state.isInPreviewMode && state.slides.length > 0) {
        const count = state.slides.length;
        dismissPreview(`${count} slide${count !== 1 ? 's' : ''} not inserted`);
    }

    state.cancelled = false;
    state.elements.welcomeState.classList.add('hidden');

    const usingBook = state.attachedBook?.status === 'ready';
    addUserMessage(usingBook ? `${content}  📄 Using: ${state.attachedBook.filename}` : content);
    messageInput.value = '';
    messageInput.style.height = 'auto';

    setProcessing(true);
    showProgressInPreviewArea('Generating content...');
    state.originalRequest = content;

    const message = state.interactivityMode
        ? { type: 'interactivity', content, interactivity: { mode: state.interactivityMode } }
        : { type: 'conversation', content };

    const sent = sendWebSocketMessage(message);
    if (!sent) connectWebSocket();
}

function handleEditSend(editInstruction) {
    const { messageInput } = state.elements;
    const cardIndex = state.editingSlideIndex;
    // The backend edits one slide of deck.slides; the title card has no layout and can't be edited.
    const slideIndex = deckSlideIndex(cardIndex);
    const currentSlide = slideIndex === null ? null : state.deck?.slides[slideIndex];
    if (!currentSlide) {
        addAIMessage('The title slide can\'t be edited — go to a content slide to edit it.');
        return;
    }

    // Show user message before preview
    const messageContent = `Edit slide ${cardIndex + 1}: ${editInstruction}`;
    const template = document.getElementById('userMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-user');
    messageEl.textContent = messageContent;
    insertBeforePreview(messageEl);

    messageInput.value = '';
    messageInput.style.height = 'auto';

    setProcessing(true);
    showProgressInPreviewArea('Updating slide...');

    const sent = sendWebSocketMessage({
        type: 'edit',
        content: editInstruction,
        edit: {
            slideIndex: slideIndex,
            currentSlide: currentSlide,
            originalRequest: state.originalRequest
        }
    });

    if (!sent) connectWebSocket();
}

function handleNewChat() {
    state.pendingRequest = null;
    state.currentSlideIndex = 0;
    setProcessing(false);
    state.conversationId = null;
    state.originalRequest = null;
    state.editingSlideIndex = null;
    state.isEditMode = false;
    state.pendingInteractivityRequest = null;
    state.pendingInteractivityMode = null;
    state.pendingActivity = null;
    exitInteractivityMode();
    detachBook();
    refreshSlideSize();

    hidePreviewArea();

    const { chatBody, welcomeState } = state.elements;
    chatBody.innerHTML = '';
    chatBody.appendChild(welcomeState);
    welcomeState.classList.remove('hidden');

    checkAndShowNPS();
}

// ============================================
// WEBSOCKET CONNECTION
// ============================================

function connectWebSocket() {
    if (state.ws && state.wsState === 'connected') return;
    if (!WS_URL) {
        console.warn('WebSocket URL not configured');
        state.wsState = 'disconnected';
        return;
    }

    state.wsState = 'connecting';
    console.log('Connecting to WebSocket:', WS_URL);

    try {
        state.ws = new WebSocket(WS_URL);
    } catch (error) {
        console.error('Failed to create WebSocket:', error);
        state.wsState = 'disconnected';
        return;
    }

    state.ws.onopen = () => {
        console.log('WebSocket connected');
        state.wsState = 'connected';
        state.reconnectAttempts = 0;
    };

    state.ws.onmessage = (event) => {
        handleWebSocketMessage(event.data);
    };

    state.ws.onclose = (event) => {
        console.log('WebSocket closed:', event.code, event.reason);
        state.wsState = 'disconnected';
        state.ws = null;

        if (event.code !== 1000 && state.reconnectAttempts < MAX_RECONNECT_ATTEMPTS) {
            state.reconnectAttempts++;
            setTimeout(connectWebSocket, 2000 * state.reconnectAttempts);
        }
    };

    state.ws.onerror = (error) => {
        console.warn('WebSocket error:', error);
    };
}

function disconnectWebSocket() {
    if (state.ws) {
        state.ws.close(1000, 'User closed');
        state.ws = null;
        state.wsState = 'disconnected';
    }
}

function sendWebSocketMessage(message) {
    if (!state.ws || state.wsState !== 'connected') {
        showError('Not connected to server. Make sure the backend is running.');
        setProcessing(false);
        return false;
    }

    try {
        const wsMessage = {
            'user-id': USER_ID,
            'channel-name': CHANNEL_NAME,
            'conversation-id': state.conversationId,
            type: message.type,
            content: message.content,
            requirements: {
                language: state.settings.language,
                level: state.settings.level,
                'native-language': 'No',
                'age-group': state.settings.ageGroup || null
            },
            // Layout hint for the AI; the backend assumes 16:9 when it's missing
            ...(state.slideSize && { 'slide-format': toSlideFormat(state.slideSize) }),
            ...(message.edit && { edit: message.edit }),
            ...(message.type === 'interactivity' && { interactivity: message.interactivity }),
            ...(state.attachedBook?.status === 'ready' && { 'book-ids': [state.attachedBook.vectorStoreId] })
        };

        state.ws.send(JSON.stringify(wsMessage));
        // From here on the book may be in the conversation's history (see handleBookChipCancel)
        if (wsMessage['book-ids']) state.attachedBook.sent = true;
        return true;
    } catch (error) {
        console.error('Failed to send WebSocket message:', error);
        showError('Failed to send message. Please try again.');
        setProcessing(false);
        return false;
    }
}

async function handleWebSocketMessage(data) {
    try {
        const message = JSON.parse(data);

        // Ignore stale responses
        if (state.staleResponses > 0) {
            if (message.type !== 'progress') state.staleResponses--;
            return;
        }

        if (message['requirements-not-met']) {
            state.conversationId = message['conversation-id'];
            hideProgress();
            addAIMessage(message['requirements-not-met']);
            setProcessing(false);
            return;
        }

        if (message.type === 'progress') {
            updateProgressInPreviewArea(message.stage || message.message);
            return;
        }

        if (message.type === 'error') {
            hideProgress();
            showError(message.message);
            setProcessing(false);
            return;
        }

        // Edit response — edit.slideIndex is an index into deck.slides
        if (message.type === 'edit' && message.edit) {
            const slideIndex = message.edit.slideIndex;
            const slide = message.edit.slide;

            if (state.deck && slide && slideIndex >= 0 && slideIndex < state.deck.slides.length) {
                state.deck.slides[slideIndex] = slide;
                syncCardsFromDeck();
                state.currentSlideIndex = cardIndexForDeckSlide(slideIndex);
                updateSlideDisplay();
                hideProgress();
                setProcessing(false);
                addAIMessage('Slide updated.');
                if (state.previewElement) {
                    setTimeout(() => {
                        state.previewElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                        if (state.isEditMode) state.elements.messageInput.focus();
                    }, 100);
                }
            } else {
                showError('Failed to apply edit. Please try again.');
                setProcessing(false);
            }
            return;
        }

        // Interactivity response
        if (message.type === 'interactivity') {
            state.conversationId = message['conversation-id'];
            hideProgress();
            exitInteractivityMode();
            showInteractivityPreview(message);
            setProcessing(false);
            return;
        }

        // Generated deck: { title, slides: [{ layout, "slide-title", … }], conversation-id }
        if (message.slides) {
            state.conversationId = message['conversation-id'];
            console.log(`[WS] Received ${message.slides.length} slide(s) from backend. Title: "${message.title}"`);
            if (message.slides.length > 0) {
                await readPresentationTheme();
                showDeckPreview({ title: message.title || null, slides: message.slides });
            } else {
                showError('No slides generated. Please try a different request.');
                setProcessing(false);
            }
            return;
        }

        if (message.error || message.message) {
            hideProgress();
            if (message.error) showError(message.error);
            else addAIMessage(message.message);
            setProcessing(false);
            return;
        }

        // Unknown
        hideProgress();
        setProcessing(false);
    } catch (error) {
        console.error('Failed to parse WebSocket message:', error);
        hideProgress();
        setProcessing(false);
    }
}

// ============================================
// DECK → PREVIEW CARDS
// ============================================

// Preview cards: the title card (when the deck has a title) followed by one card per AI slide.
function syncCardsFromDeck() {
    const { deck } = state;
    state.slides = deck
        ? [
            ...(deck.title ? [{ kind: 'title', title: deck.title }] : []),
            ...deck.slides.map((slide) => ({ kind: 'layout', slide })),
        ]
        : [];
}

// Card index → index into deck.slides (null for the title card).
function deckSlideIndex(cardIndex) {
    if (!state.deck) return null;
    const offset = state.deck.title ? 1 : 0;
    return cardIndex >= offset ? cardIndex - offset : null;
}

function cardIndexForDeckSlide(slideIndex) {
    return slideIndex + (state.deck?.title ? 1 : 0);
}

function showDeckPreview(deck) {
    state.deck = deck;
    syncCardsFromDeck();
    showSlidePreview(state.slides, deck.title || 'Generated Content');
}

// ============================================
// UI - Messages
// ============================================

function addUserMessage(content) {
    const template = document.getElementById('userMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-user');
    messageEl.textContent = content;
    appendToChatBody(messageEl);
}

function addAIMessage(displayContent) {
    const template = document.getElementById('aiMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-ai');
    messageEl.textContent = displayContent;
    appendToChatBody(messageEl);
}

// ============================================
// UI - Progress
// ============================================

function hideProgress() {
    if (state.progressElement && state.progressElement.parentNode) {
        state.progressElement.remove();
    }
    state.progressElement = null;
}

function cancelGeneration() {
    state.cancelled = true;
    state.conversationId = null;
    state.staleResponses++;
    hideProgress();
    setProcessing(false);
    addAIMessage('Generation cancelled.');
}

function showProgressInPreviewArea(status) {
    const template = document.getElementById('progressMessageTemplate');
    const clone = template.content.cloneNode(true);
    const progressEl = clone.querySelector('.message-progress');
    progressEl.querySelector('.progress-status').textContent = status;

    const cancelBtn = progressEl.querySelector('.progress-cancel-btn');
    if (cancelBtn) cancelBtn.addEventListener('click', cancelGeneration);

    state.progressElement = progressEl;
    appendToChatBody(progressEl);
}

function updateProgressInPreviewArea(status) {
    if (!state.progressElement) {
        showProgressInPreviewArea(status);
        return;
    }
    state.progressElement.querySelector('.progress-status').textContent = status;
}

function hidePreviewArea() {
    if (state.previewElement && state.previewElement.parentNode) {
        state.previewElement.remove();
    }
    state.deck = null;
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
}

function dismissPreview(message) {
    if (state.previewElement && state.previewElement.parentNode) {
        state.previewElement.remove();
    }
    addAIMessage(message);
    state.deck = null;
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
    state.conversationId = null;
    state.pendingActivity = null;
}

// ============================================
// UI - Slide Preview
// ============================================

function showSlidePreview(slides, summary) {
    if (state.progressElement && state.progressElement.parentNode) {
        state.progressElement.remove();
        state.progressElement = null;
    }

    state.slides = slides;
    state.currentSlideIndex = 0;
    state.isInPreviewMode = true;
    setProcessing(false);

    const template = document.getElementById('previewContainerTemplate');
    const clone = template.content.cloneNode(true);
    const previewEl = clone.querySelector('.preview-container');

    previewEl.querySelector('.preview-title-text').textContent = summary || 'Preview';
    previewEl.querySelector('.preview-count').textContent = `${slides.length} slides`;

    state.previewElement = previewEl;
    appendToChatBody(previewEl);
    updateSlideDisplay();
    setupPreviewNavigation(previewEl);

    setTimeout(() => {
        previewEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 100);
}

function updateSlideDisplay() {
    if (!state.previewElement || !state.slides.length) return;

    const slide = state.slides[state.currentSlideIndex];
    const previewEl = state.previewElement;

    previewEl.querySelector('.slide-counter-text').textContent =
        `Slide ${state.currentSlideIndex + 1} of ${state.slides.length}`;

    const miniContainer = previewEl.querySelector('.slide-mini-container');
    const textCard = previewEl.querySelector('.slide-card');
    const editBtn = previewEl.querySelector('#navEditBtn');

    if (slide.kind) {
        // Layout slide / deck title — mini slide preview
        previewEl.querySelector('.slide-type-badge').textContent = cardLabel(slide);
        renderSlideCard(miniContainer, slide);
        miniContainer.classList.remove('hidden');
        textCard.classList.add('hidden');
        // The title slide has no layout for the backend to edit
        editBtn.classList.toggle('hidden', slide.kind === 'title');
    } else {
        // Interactivity question — plain text card
        miniContainer.classList.add('hidden');
        textCard.classList.remove('hidden');
        previewEl.querySelector('.slide-type-badge').textContent = slide.type || 'Content';
        previewEl.querySelector('.slide-card-title').textContent = slide.title || '';
        previewEl.querySelector('.slide-card-subtitle').textContent = '';
        previewEl.querySelector('.slide-card-content').textContent = slide.content || '';
        previewEl.querySelector('.slide-card-example').classList.add('hidden');
    }

    const backBtn = previewEl.querySelector('#navBackBtn');
    const nextBtn = previewEl.querySelector('#navNextBtn');
    backBtn.disabled = state.currentSlideIndex === 0;

    const slideCount = state.slides.length;
    if (state.currentSlideIndex === state.slides.length - 1) {
        nextBtn.innerHTML = `
            <span class="material-icons" style="font-size: 16px;">playlist_add</span>
            Insert ${slideCount}
            <span class="nav-shortcut">Enter</span>
        `;
    } else {
        nextBtn.innerHTML = `
            <span class="material-icons" style="font-size: 16px;">arrow_forward</span>
            Next
            <span class="nav-shortcut">Enter</span>
        `;
    }
}

function setupPreviewNavigation(previewEl) {
    previewEl.querySelector('#navBackBtn').addEventListener('click', navigateBack);
    previewEl.querySelector('#navSkipBtn').addEventListener('click', removeSlide);
    previewEl.querySelector('#navEditBtn').addEventListener('click', editSlide);
    previewEl.querySelector('#navNextBtn').addEventListener('click', navigateNext);
    previewEl.querySelector('#previewCancelBtn').addEventListener('click', cancelPreview);

    // Interactivity activities (quiz/sentence-ordering/guess-the-word/true-false) insert as a
    // single slide built from state.pendingActivity — per-question Edit/Remove has nothing to act on.
    if (state.pendingActivity) {
        previewEl.querySelector('#navSkipBtn').classList.add('hidden');
        previewEl.querySelector('#navEditBtn').classList.add('hidden');
    }
}

function cancelPreview() {
    dismissPreview('Preview cancelled.');
}

function navigateBack() {
    if (state.currentSlideIndex > 0) {
        state.currentSlideIndex--;
        if (state.isEditMode) {
            state.editingSlideIndex = state.currentSlideIndex;
            updateEditBadgeSlideNumber(state.currentSlideIndex + 1);
        }
        updateSlideDisplay();
    }
}

function navigateNext() {
    if (state.currentSlideIndex < state.slides.length - 1) {
        state.currentSlideIndex++;
        if (state.isEditMode) {
            state.editingSlideIndex = state.currentSlideIndex;
            updateEditBadgeSlideNumber(state.currentSlideIndex + 1);
        }
        updateSlideDisplay();
    } else {
        insertAllSlides();
    }
}

function removeSlide() {
    if (state.pendingActivity) return;

    if (state.slides.length <= 1) {
        dismissPreview('All slides removed.');
        return;
    }

    // Remove from the deck (source of truth), then rebuild the cards
    const slideIndex = deckSlideIndex(state.currentSlideIndex);
    if (slideIndex === null) state.deck.title = null;   // the title card: insert without a title slide
    else state.deck.slides.splice(slideIndex, 1);
    syncCardsFromDeck();
    if (state.currentSlideIndex >= state.slides.length) {
        state.currentSlideIndex = state.slides.length - 1;
    }
    if (state.isEditMode) {
        state.editingSlideIndex = state.currentSlideIndex;
        updateEditBadgeSlideNumber(state.currentSlideIndex + 1);
    }
    updateSlideDisplay();
}

function editSlide() {
    if (state.pendingActivity) return;
    if (deckSlideIndex(state.currentSlideIndex) === null) return; // title card — nothing for the backend to edit

    const { messageInput } = state.elements;
    state.isEditMode = true;
    state.editingSlideIndex = state.currentSlideIndex;
    messageInput.placeholder = `Describe what to change on slide ${state.currentSlideIndex + 1}...`;
    messageInput.value = '';
    showEditBadge(state.currentSlideIndex + 1);
    messageInput.focus();
}

function exitEditMode() {
    const { messageInput } = state.elements;
    state.isEditMode = false;
    state.editingSlideIndex = null;
    messageInput.placeholder = 'Type your request...';
    removeEditBadge();
}

function showEditBadge(slideNumber) {
    const { editBadge, editBadgeSlideNum } = state.elements;
    if (editBadge && editBadgeSlideNum) {
        editBadgeSlideNum.textContent = slideNumber;
        editBadge.classList.remove('hidden');
    }
}

function updateEditBadgeSlideNumber(slideNumber) {
    const { editBadgeSlideNum, messageInput } = state.elements;
    if (editBadgeSlideNum) editBadgeSlideNum.textContent = slideNumber;
    if (messageInput) messageInput.placeholder = `Describe what to change on slide ${slideNumber}...`;
}

function removeEditBadge() {
    const { editBadge } = state.elements;
    if (editBadge) editBadge.classList.add('hidden');
}

function enterInteractivityMode(mode) {
    const { interactivityChip, interactivityChipIcon, interactivityChipLabel, messageInput } = state.elements;
    state.interactivityMode = mode;
    const meta = INTERACTIVITY_MODE_LABELS[mode] || INTERACTIVITY_MODE_LABELS['multiple-choice'];
    if (interactivityChip) interactivityChip.classList.remove('hidden');
    if (interactivityChipIcon) interactivityChipIcon.textContent = meta.icon;
    if (interactivityChipLabel) interactivityChipLabel.textContent = meta.label;
    if (messageInput) messageInput.placeholder = meta.placeholder;
}

function exitInteractivityMode() {
    const { interactivityChip, messageInput } = state.elements;
    state.interactivityMode = null;
    if (interactivityChip) interactivityChip.classList.add('hidden');
    if (messageInput) messageInput.placeholder = 'Type your request...';
}

// ============================================
// ATTACHED BOOK (file_search grounding)
// ============================================

const BOOK_MAX_SIZE_BYTES = 20 * 1024 * 1024; // matches the backend's own limit
const BOOK_POLL_INTERVAL_MS = 3000;
const BOOK_POLL_MAX_ATTEMPTS = 40; // ~2 minutes

function handleBookFileSelected(file) {
    if (file.type !== 'application/pdf') {
        showError('Only PDF files are supported.');
        resetBookFileInput();
        return;
    }
    if (file.size > BOOK_MAX_SIZE_BYTES) {
        showError('That file is too large (20MB max).');
        resetBookFileInput();
        return;
    }
    uploadBook(file);
}

function resetBookFileInput() {
    const { bookFileInput } = state.elements;
    if (bookFileInput) bookFileInput.value = '';
}

async function uploadBook(file) {
    state.attachedBook = { filename: file.name, status: 'uploading' };
    renderBookChip();

    const formData = new FormData();
    formData.append('file', file);
    formData.append('book-name', file.name);

    try {
        const res = await fetch(`${API_URL}/library/upload`, { method: 'POST', body: formData });
        if (!res.ok) throw new Error('Upload failed');
        const data = await res.json();

        state.attachedBook = {
            vectorStoreId: data['vector-store-id'],
            filename: file.name,
            status: 'processing'
        };
        renderBookChip();
        startBookStatusPolling();
    } catch (error) {
        console.error('[Library] Upload failed:', error);
        showError('Failed to upload the document. Please try again.');
        detachBook();
    } finally {
        resetBookFileInput();
    }
}

function startBookStatusPolling() {
    state.bookPollAttempts = 0;
    clearInterval(state.bookPollTimer);
    state.bookPollTimer = setInterval(pollBookStatus, BOOK_POLL_INTERVAL_MS);
}

async function pollBookStatus() {
    if (!state.attachedBook || !state.attachedBook.vectorStoreId) return;

    state.bookPollAttempts++;
    if (state.bookPollAttempts > BOOK_POLL_MAX_ATTEMPTS) {
        showError('Your document is taking longer than expected to process. Please try again.');
        detachBook();
        return;
    }

    try {
        const res = await fetch(`${API_URL}/library/book-status/${state.attachedBook.vectorStoreId}`);
        if (!res.ok) throw new Error('Status check failed');
        const data = await res.json();

        if (data.status === 'completed') {
            clearInterval(state.bookPollTimer);
            state.attachedBook.status = 'ready';
            renderBookChip();
            addInfoMessage(`"${state.attachedBook.filename}" is ready — your next message will use it. (Free tier: removed after a few days.)`);
        } else if (data.status === 'expired' || (data['file-counts'] && data['file-counts'].failed > 0)) {
            throw new Error('Processing failed');
        }
        // otherwise still "in_progress" — keep polling
    } catch (error) {
        console.error('[Library] Status check failed:', error);
        showError('Failed to process the attached document. Please try again.');
        detachBook();
    }
}

// A book that was sent may already be in the conversation's history (file_search passages travel
// with previous_response_id), so removing it mid-conversation would let the next answer keep using
// it or mix it with a newly attached book. Ask first; on yes, the conversation starts over.
// Only a conversation the next message would continue counts: with a preview open, the next
// message either edits a slide (no history) or dismisses the preview, which resets it anyway.
function handleBookChipCancel() {
    const book = state.attachedBook;
    if (!book || !book.sent || !state.conversationId || state.isInPreviewMode) {
        detachBook(); // nothing in the history the next message would carry on with
        return;
    }
    if (state.bookChangeConfirmEl) return; // already asking

    const template = document.getElementById('aiMessageTemplate');
    const messageEl = template.content.cloneNode(true).querySelector('.message-ai');
    messageEl.innerHTML = `
        <div>📄 Changing the document starts a new conversation — the previous history is lost.</div>
        <div style="margin-top: 6px;">Continue?</div>
        <div class="interactivity-confirm-buttons">
            <button class="interactivity-confirm-btn interactivity-confirm-yes">Remove document</button>
            <button class="interactivity-confirm-btn interactivity-confirm-no">Keep it</button>
        </div>
    `;
    messageEl.querySelector('.interactivity-confirm-yes').addEventListener('click', () => {
        if (state.isProcessing) return; // a request sent after the bubble appeared — answer it once that's done
        state.bookChangeConfirmEl = null;
        if (state.attachedBook !== book) { messageEl.remove(); return; } // already changed (e.g. New chat)
        state.conversationId = null;
        detachBook();
        messageEl.textContent = `📄 "${book.filename}" removed — your next message starts a new conversation, so include the topic again.`;
    });
    messageEl.querySelector('.interactivity-confirm-no').addEventListener('click', () => {
        state.bookChangeConfirmEl = null;
        messageEl.remove();
    });

    state.bookChangeConfirmEl = messageEl;
    appendToChatBody(messageEl);
}

function detachBook() {
    if (state.bookChangeConfirmEl) {
        state.bookChangeConfirmEl.remove();
        state.bookChangeConfirmEl = null;
    }
    clearInterval(state.bookPollTimer);
    state.bookPollTimer = null;
    state.bookPollAttempts = 0;
    state.attachedBook = null;
    renderBookChip();
    resetBookFileInput();
}

function renderBookChip() {
    const { bookChip, bookChipIcon, bookChipLabel, bookChipStatus, attachBookBtn } = state.elements;
    if (!bookChip) return;

    const book = state.attachedBook;
    updateInputState();
    if (!book) {
        bookChip.classList.add('hidden');
        bookChip.classList.remove('book-chip-processing', 'book-chip-blocked');
        if (attachBookBtn) attachBookBtn.disabled = false;
        return;
    }

    bookChip.classList.remove('hidden');
    if (attachBookBtn) attachBookBtn.disabled = true; // one book at a time
    bookChipLabel.textContent = book.filename;

    const isProcessing = book.status === 'uploading' || book.status === 'processing';
    bookChip.classList.toggle('book-chip-processing', isProcessing);
    if (!isProcessing) bookChip.classList.remove('book-chip-blocked');

    if (book.status === 'uploading') {
        bookChipIcon.textContent = 'description';
        bookChipStatus.textContent = 'Uploading…';
    } else if (book.status === 'processing') {
        bookChipIcon.textContent = 'description';
        bookChipStatus.textContent = 'Processing…';
    } else if (book.status === 'ready') {
        bookChipIcon.textContent = 'check_circle';
        bookChipStatus.textContent = 'Ready';
    }
}

function addInfoMessage(text) {
    const template = document.getElementById('infoMessageTemplate');
    const clone = template.content.cloneNode(true);
    const infoEl = clone.querySelector('.message-info');
    infoEl.querySelector('.info-text').textContent = text;
    appendToChatBody(infoEl);
}

// ── Interactivity preview ─────────────────────────────────────────────────────

function buildInteractivityPreviewSlides(activity, questions) {
    switch (activity.mode) {
        case 'sentence-ordering':
            return questions.map((q, i) => ({
                title: `${i + 1}. Arrange the words`,
                content: (q.sentence || []).join(' ') + (q.hint ? `\n\n💡 ${q.hint}` : ''),
                type: 'Sentence Ordering'
            }));
        case 'guess-the-word':
            return questions.map((q, i) => ({
                title: `${i + 1}. ${q.definition}`,
                content: `Answer: ${q.word}` + (q.hint ? `\n💡 ${q.hint}` : ''),
                type: 'Guess the Word'
            }));
        case 'true-false':
            return questions.map((q, i) => ({
                title: `${i + 1}. ${q.statement}`,
                content: `Answer: ${q.answer ? 'True' : 'False'}` + (q.explanation ? `\n\n→ ${q.explanation}` : ''),
                type: 'True or False'
            }));
        default:
            return questions.map((q, i) => ({
                title: `${i + 1}. ${q.question}`,
                content: q.options.map((opt, oi) =>
                    `${oi === q.correct ? '✓' : '   '} ${String.fromCharCode(65 + oi)}) ${opt}`
                ).join('\n') + (q.explanation ? `\n\n→ ${q.explanation}` : ''),
                type: 'Quiz'
            }));
    }
}

function showInteractivityPreview(activity) {
    state.pendingActivity = activity;
    state.deck = null;

    const questions = activity.questions || [];
    const slides = buildInteractivityPreviewSlides(activity, questions);

    showSlidePreview(slides, activity.title || (INTERACTIVITY_MODE_LABELS[activity.mode] || {}).label || 'Interactive Activity');
}

async function handleAddInteractivityToSlide(activity) {
    hidePreviewArea();
    state.pendingActivity = null;

    setProcessing(true);
    showProgressInPreviewArea('Building activity slide...');

    try {
        const base64 = await buildInteractivityPptxBase64(activity, await refreshSlideSize());

        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            presentation.slides.load('items/id');
            await context.sync();

            const items = presentation.slides.items;
            const lastSlideId = items.length > 0 ? items[items.length - 1].id : undefined;

            presentation.insertSlidesFromBase64(base64, {
                formatting: 'UseDestinationTheme',
                targetSlideId: lastSlideId,
            });
            await context.sync();
        });

        hideProgress();
        setProcessing(false);
        state.conversationId = null;
        showSuccess((INTERACTIVITY_MODE_LABELS[activity.mode] || {}).successMessage
            || 'Activity slide added! Students can scan the QR code to play.');
    } catch (error) {
        console.error('[Interactivity] Failed to insert slide:', error);
        hideProgress();
        setProcessing(false);
        showError('Failed to add activity slide. Please try again.');
    }
}

// ── Interactivity keyword detection ──────────────────────────────────────────

function detectInteractivityMode(text) {
    const lower = text.toLowerCase();
    const match = INTERACTIVITY_KEYWORDS.find(k => lower.includes(k.keyword));
    return match ? match.mode : null;
}

function showInteractivityConfirmation(originalMessage, mode) {
    state.pendingInteractivityRequest = originalMessage;
    state.pendingInteractivityMode = mode;

    const template = document.getElementById('aiMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-ai');

    messageEl.innerHTML = `
        <div>🎮 This looks like an interactive activity.</div>
        <div style="margin-top: 6px;">How would you like to proceed?</div>
        <div class="interactivity-confirm-buttons">
            <button class="interactivity-confirm-btn interactivity-confirm-yes">Make it interactive</button>
            <button class="interactivity-confirm-btn interactivity-confirm-no">Create slides</button>
        </div>
    `;

    messageEl.querySelector('.interactivity-confirm-yes').addEventListener('click', (e) => handleInteractivityConfirmYes(e, messageEl));
    messageEl.querySelector('.interactivity-confirm-no').addEventListener('click', (e) => handleInteractivityConfirmNo(e, messageEl));

    appendToChatBody(messageEl);
}

function handleInteractivityConfirmYes(e, messageEl) {
    // A book attached after this bubble appeared — keep the bubble so it can be clicked once ready
    if (isBookPending()) { flagBookPending(); return; }
    const mode = state.pendingInteractivityMode || 'multiple-choice';
    const meta = INTERACTIVITY_MODE_LABELS[mode] || INTERACTIVITY_MODE_LABELS['multiple-choice'];
    messageEl.textContent = `🎮 ${meta.label}`;
    const pending = state.pendingInteractivityRequest;
    state.pendingInteractivityRequest = null;
    state.pendingInteractivityMode = null;
    enterInteractivityMode(mode);
    if (pending) triggerSendWithContent(pending);
}

function handleInteractivityConfirmNo(e, messageEl) {
    if (isBookPending()) { flagBookPending(); return; }
    messageEl.textContent = '📄 Creating slides...';
    const pending = state.pendingInteractivityRequest;
    state.pendingInteractivityRequest = null;
    state.pendingInteractivityMode = null;
    if (pending) triggerSendWithContent(pending);
}

function triggerSendWithContent(content) {
    state.cancelled = false;
    setProcessing(true);
    showProgressInPreviewArea('Generating content...');
    state.originalRequest = content;

    const message = state.interactivityMode
        ? { type: 'interactivity', content, interactivity: { mode: state.interactivityMode } }
        : { type: 'conversation', content };

    const sent = sendWebSocketMessage(message);
    if (!sent) connectWebSocket();
}

// ── Commands autocomplete ─────────────────────────────────────────────────────

let highlightedCommandIndex = -1;

function showCommandsAutocomplete(query) {
    const { commandsAutocomplete } = state.elements;
    if (!commandsAutocomplete) return;

    const matches = COMMANDS.filter(c => c.command.startsWith(query.toLowerCase()));

    if (matches.length === 0) {
        hideCommandsAutocomplete();
        return;
    }

    highlightedCommandIndex = -1;
    commandsAutocomplete.innerHTML = matches.map((c, i) => `
        <div class="command-item" data-index="${i}" data-command="${c.command}" data-mode="${c.mode}">
            <span class="material-icons">${c.icon}</span>
            <span class="command-item-name">${c.command}</span>
            <span class="command-item-desc">${c.description}</span>
        </div>
    `).join('');

    commandsAutocomplete.querySelectorAll('.command-item').forEach(item => {
        item.addEventListener('mousedown', (e) => {
            e.preventDefault();
            selectCommand(item.dataset.mode);
        });
    });

    commandsAutocomplete.classList.remove('hidden');
}

function hideCommandsAutocomplete() {
    const { commandsAutocomplete } = state.elements;
    if (commandsAutocomplete) commandsAutocomplete.classList.add('hidden');
    highlightedCommandIndex = -1;
}

function selectCommand(mode) {
    const { messageInput } = state.elements;
    if (messageInput) messageInput.value = '';
    hideCommandsAutocomplete();
    enterInteractivityMode(mode);
    if (messageInput) messageInput.focus();
}

function handleCommandAutocompleteKeydown(e) {
    const { commandsAutocomplete } = state.elements;
    if (!commandsAutocomplete || commandsAutocomplete.classList.contains('hidden')) return;

    const items = commandsAutocomplete.querySelectorAll('.command-item');
    if (items.length === 0) return;

    if (e.key === 'ArrowDown') {
        e.preventDefault();
        highlightedCommandIndex = (highlightedCommandIndex + 1) % items.length;
    } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        highlightedCommandIndex = (highlightedCommandIndex - 1 + items.length) % items.length;
    } else if (e.key === 'Enter' && highlightedCommandIndex >= 0) {
        e.preventDefault();
        selectCommand(items[highlightedCommandIndex].dataset.mode);
        return;
    } else if (e.key === 'Escape') {
        hideCommandsAutocomplete();
        return;
    } else {
        return;
    }

    items.forEach((item, i) => item.classList.toggle('highlighted', i === highlightedCommandIndex));
}

// ============================================
// POWERPOINT SLIDE INSERTION
// ============================================

// One path on desktop and web: build a .pptx at the deck's real size with the layout drawers
// (src/taskpane/slides/), then insertSlidesFromBase64 with the destination theme.
async function insertAllSlides() {
    if (state.pendingActivity) {
        await handleAddInteractivityToSlide(state.pendingActivity);
        return;
    }

    const deck = state.deck;
    if (!deck || (!deck.title && deck.slides.length === 0)) {
        showError('No slides to insert.');
        return;
    }

    await readPresentationTheme();

    if (state.isEditMode) exitEditMode();
    hidePreviewArea();

    setProcessing(true);
    showProgressInPreviewArea('Building slides...');

    try {
        const size = (await refreshSlideSize()) || { w: 13.333, h: 7.5 };
        const { base64, slideCount, splits, skipped } = await buildDeckBase64(deck, size, SLIDE_THEME.fonts);
        console.log(`[Insert] ${slideCount} slide(s) at ${size.w.toFixed(2)}×${size.h.toFixed(2)}in — ${splits} split, ${skipped} skipped.`);

        updateProgressInPreviewArea('Inserting slides...');
        await insertDeckBase64(base64);

        hideProgress();
        const splitNote = splits > 0 ? ` (${splits} split to fit)` : '';
        showSuccess(`${slideCount} slide${slideCount !== 1 ? 's' : ''} inserted successfully${splitNote}`);
        const fbStats = getFeedbackStats();
        saveFeedbackStats({ ...fbStats, insertCount: fbStats.insertCount + 1 });
        setProcessing(false);
        state.conversationId = null;
    } catch (error) {
        console.error('[Insert] Error inserting slides:', error);
        hideProgress();
        showError(`Failed to insert slides: ${error.message}`);
        setProcessing(false);
    }
}

async function buildInteractivityPptxBase64(activity, size) {
    // Designed on a 10×5.625in (720×405pt) canvas; scaled uniformly and centred on the deck's real size.
    const W = size?.w || 10;
    const H = size?.h || 5.625;
    const scale = Math.min(W / 10, H / 5.625);
    const ox = (W - 10 * scale) / 2;
    const oy = (H - 5.625 * scale) / 2;
    const sz = (v) => +((v / 72) * scale).toFixed(4);   // Office JS points → PptxGenJS inches
    const px = (v) => +(ox + sz(v)).toFixed(4);
    const py = (v) => +(oy + sz(v)).toFixed(4);
    const fs = (v) => Math.round(v * scale);
    const c = (hex) => hex.slice(1);
    const meta = INTERACTIVITY_MODE_LABELS[activity.mode] || INTERACTIVITY_MODE_LABELS['multiple-choice'];

    const gameUrl = `${GAME_BASE_URL}/${activity['activity-id']}`;
    const qrDataUrl = await QRCode.toDataURL(gameUrl, { width: 300, margin: 1 });

    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DEST', width: W, height: H });
    pptx.layout = 'DEST';
    const slide = pptx.addSlide();

    // Left — QR code
    slide.addImage({
        data: qrDataUrl,
        x: px(50), y: py(90),
        w: sz(260), h: sz(260),
    });

    // Left label — "Scan to play!"
    slide.addText('Scan to play!', {
        x: px(50), y: py(360),
        w: sz(260), h: sz(30),
        fontSize: fs(14),
        bold: true,
        fontFace: SLIDE_THEME.fonts.body,
        color: c(SLIDE_THEME.colors.subtitle),
        align: 'center',
    });

    // Right — Activity title
    slide.addText(activity.title || meta.label, {
        x: px(360), y: py(90),
        w: sz(310), h: sz(70),
        fontSize: fs(28),
        bold: true,
        fontFace: SLIDE_THEME.fonts.heading,
        color: c(SLIDE_THEME.colors.title),
        wrap: true,
    });

    // Right — Question count
    const questionCount = activity['question-count'] || (activity.questions || []).length;
    slide.addText(`🎮  ${questionCount} ${meta.countNoun}`, {
        x: px(360), y: py(170),
        w: sz(310), h: sz(30),
        fontSize: fs(16),
        fontFace: SLIDE_THEME.fonts.body,
        color: c(SLIDE_THEME.colors.subtitle),
    });

    // Right — Instructions
    slide.addText(`How to join:\n1. Open your phone camera\n2. Scan the QR code\n3. ${meta.joinStep3}`, {
        x: px(360), y: py(230),
        w: sz(310), h: sz(160),
        fontSize: fs(16),
        fontFace: SLIDE_THEME.fonts.body,
        color: c(SLIDE_THEME.colors.content),
        wrap: true,
    });

    return await pptx.write('base64');
}

// ============================================
// UI - Success/Error
// ============================================

function showSuccess(message) {
    hidePreviewArea();
    const template = document.getElementById('successMessageTemplate');
    const clone = template.content.cloneNode(true);
    const successEl = clone.querySelector('.message-success');
    successEl.querySelector('.success-text').textContent = message;

    const thumbsRow = document.createElement('div');
    thumbsRow.className = 'feedback-thumbs';
    thumbsRow.innerHTML = `
        <span class="feedback-prompt">Was this helpful?</span>
        <button class="thumb-btn" data-val="up" title="Yes">👍</button>
        <button class="thumb-btn" data-val="down" title="No">👎</button>
    `;
    successEl.appendChild(thumbsRow);
    thumbsRow.querySelectorAll('.thumb-btn').forEach(btn => {
        btn.addEventListener('click', () => handleThumbsFeedback(btn.dataset.val, thumbsRow));
    });

    appendToChatBody(successEl);
}

function showError(message) {
    hidePreviewArea();
    const template = document.getElementById('errorMessageTemplate');
    const clone = template.content.cloneNode(true);
    const errorEl = clone.querySelector('.message-error');
    errorEl.querySelector('.error-text').textContent = message;

    const feedbackLink = document.createElement('button');
    feedbackLink.className = 'feedback-error-link';
    feedbackLink.textContent = 'Tell us what happened →';
    feedbackLink.addEventListener('click', () => openFeedbackModal({ errorMessage: message }));
    errorEl.appendChild(feedbackLink);

    appendToChatBody(errorEl);
}

// ============================================
// KEYBOARD NAVIGATION
// ============================================

function flashButton(btnId) {
    if (!state.previewElement) return;
    const btn = state.previewElement.querySelector(`#${btnId}`);
    if (!btn || btn.disabled) return;
    btn.classList.add('nav-btn-pressed');
    setTimeout(() => btn.classList.remove('nav-btn-pressed'), 150);
}

function handleGlobalKeydown(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        const hasText = state.elements.messageInput.value.trim().length > 0;
        if (hasText && !state.isProcessing) {
            handleSend();
        } else if (state.isInPreviewMode && !state.isProcessing) {
            flashButton('navNextBtn');
            navigateNext();
        }
        return;
    }

    if (e.key === 'Escape') {
        if (state.isEditMode) {
            e.preventDefault();
            exitEditMode();
            return;
        }
    }

    if (e.key.toLowerCase() === 'q' && document.activeElement !== state.elements.messageInput) {
        if (state.isProcessing) { e.preventDefault(); cancelGeneration(); return; }
        if (state.isInPreviewMode) { e.preventDefault(); cancelPreview(); return; }
    }

    if (!state.isInPreviewMode || state.isProcessing) return;
    if (document.activeElement === state.elements.messageInput) return;

    switch (e.key.toLowerCase()) {
        case 'r': if (!state.pendingActivity) { e.preventDefault(); flashButton('navSkipBtn'); removeSlide(); } break;
        case 'e': if (!state.pendingActivity) { e.preventDefault(); flashButton('navEditBtn'); editSlide(); } break;
        case 'a': e.preventDefault(); insertAllSlides(); break;
        case 'escape': e.preventDefault(); state.elements.messageInput.focus(); break;
        case 'arrowleft':
        case 'backspace':
            if (state.currentSlideIndex > 0) { e.preventDefault(); flashButton('navBackBtn'); navigateBack(); }
            break;
        case 'arrowright':
            if (state.currentSlideIndex < state.slides.length - 1) { e.preventDefault(); flashButton('navNextBtn'); navigateNext(); }
            break;
    }
}

// ============================================
// SETTINGS
// ============================================

function openSettingsModal() {
    state.elements.settingsModal.classList.remove('hidden');
}

function closeSettingsModal() {
    if (!state.settingsConfirmed) return;
    state.elements.settingsModal.classList.add('hidden');
    loadSettingsToForm();
}

function getSettingsStorageKey() {
    if (state.currentFilePath) {
        const sanitizedPath = state.currentFilePath.replace(/[^a-zA-Z0-9]/g, '_');
        return `teachersCenterSettings_${sanitizedPath}`;
    }
    return 'teachersCenterSettings_default';
}

function loadSettings() {
    try {
        state.currentFilePath = Office.context.document?.url || null;
    } catch (error) {
        state.currentFilePath = null;
    }

    const storageKey = getSettingsStorageKey();
    const savedSettings = localStorage.getItem(storageKey);

    if (savedSettings) {
        state.settings = JSON.parse(savedSettings);
        state.settingsConfirmed = true;
    } else {
        state.settings = { language: 'English', level: 'B1', ageGroup: '' };
        state.settingsConfirmed = false;
        setTimeout(() => openSettingsModal(), 100);
    }

    loadSettingsToForm();
    updateContextBadge();

    // Returning users (settings already confirmed) never hit saveSettings(),
    // so show What's New here instead — first-run users get it from saveSettings().
    if (state.settingsConfirmed) {
        setTimeout(() => maybeShowWhatsNewModal(), 200);
    }
}

function loadSettingsToForm() {
    const { settingsLanguage, settingsLevel, settingsAgeGroup } = state.elements;
    settingsLanguage.value = state.settings.language;
    settingsLevel.value = state.settings.level;
    settingsAgeGroup.value = state.settings.ageGroup;
}

function saveSettings() {
    const { settingsLanguage, settingsLevel, settingsAgeGroup } = state.elements;

    state.settings = {
        language: settingsLanguage.value,
        level: settingsLevel.value,
        ageGroup: settingsAgeGroup.value
    };

    const storageKey = getSettingsStorageKey();
    localStorage.setItem(storageKey, JSON.stringify(state.settings));
    state.settingsConfirmed = true;
    updateContextBadge();
    closeSettingsModal();
    setTimeout(() => maybeShowWhatsNewModal(), 200);
}

function updateContextBadge() {
    const { contextBadge } = state.elements;
    if (contextBadge) {
        contextBadge.textContent = `${state.settings.level} ${state.settings.language}`;
    }
}

// ============================================
// WHAT'S NEW
// ============================================

const WHATS_NEW_STORAGE_KEY = 'teachersCenterWhatsNewSeen_interactivity';

function maybeShowWhatsNewModal() {
    if (localStorage.getItem(WHATS_NEW_STORAGE_KEY)) return;
    state.elements.whatsNewModal.classList.remove('hidden');
}

function closeWhatsNewModal() {
    if (state.elements.whatsNewDontShowAgain.checked) {
        localStorage.setItem(WHATS_NEW_STORAGE_KEY, 'true');
    }
    state.elements.whatsNewModal.classList.add('hidden');
}

// ============================================
// FEEDBACK
// ============================================

async function sendFeedback(data) {
    try {
        await fetch(`${API_URL}/feedback`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                ...data,
                settings: state.settings,
                timestamp: new Date().toISOString()
            })
        });
    } catch (_) { /* silent — feedback is non-critical */ }
}

function getFeedbackStats() {
    try {
        return JSON.parse(localStorage.getItem('teachersCenterFeedbackStats'))
            || { insertCount: 0, lastPromptedAt: null, neverAsk: false };
    } catch (_) {
        return { insertCount: 0, lastPromptedAt: null, neverAsk: false };
    }
}

function saveFeedbackStats(stats) {
    localStorage.setItem('teachersCenterFeedbackStats', JSON.stringify(stats));
}

function handleThumbsFeedback(val, container) {
    if (val === 'up') {
        sendFeedback({ type: 'success-rating', rating: 'up' });
        container.innerHTML = '<span class="feedback-thanks">Thanks! 🙏</span>';
    } else {
        container.innerHTML = `
            <textarea class="feedback-inline-text" placeholder="What could be better? (optional)" rows="2"></textarea>
            <div class="feedback-inline-actions">
                <button class="feedback-send-btn">Send</button>
                <button class="feedback-skip-btn">Skip</button>
            </div>`;
        container.querySelector('.feedback-send-btn').addEventListener('click', () => {
            const comment = container.querySelector('.feedback-inline-text').value;
            sendFeedback({ type: 'success-rating', rating: 'down', comment });
            container.innerHTML = '<span class="feedback-thanks">Thanks for the feedback! 🙏</span>';
        });
        container.querySelector('.feedback-skip-btn').addEventListener('click', () => {
            container.remove();
        });
    }
}

function openFeedbackModal(context) {
    selectedStarRating = 0;
    state.elements.feedbackComment.value = '';
    state.elements.submitFeedbackBtn.disabled = true;
    state.elements.starRating.querySelectorAll('.star').forEach(s => s.classList.remove('active'));

    if (context?.errorMessage) {
        state.elements.feedbackErrorContext.textContent = `Error context: ${context.errorMessage}`;
        state.elements.feedbackErrorContext.classList.remove('hidden');
        state.elements.feedbackErrorContext.dataset.errorMessage = context.errorMessage;
    } else {
        state.elements.feedbackErrorContext.classList.add('hidden');
        delete state.elements.feedbackErrorContext.dataset.errorMessage;
    }
    state.elements.feedbackModal.classList.remove('hidden');
}

function closeFeedbackModal() {
    state.elements.feedbackModal.classList.add('hidden');
}

function submitFeedbackModal() {
    const comment = state.elements.feedbackComment.value.trim();
    const errorMessage = state.elements.feedbackErrorContext.dataset.errorMessage;
    sendFeedback({
        type: errorMessage ? 'error-feedback' : 'manual-feedback',
        rating: selectedStarRating,
        comment,
        ...(errorMessage && { errorContext: errorMessage })
    });
    closeFeedbackModal();
}

function checkAndShowNPS() {
    const stats = getFeedbackStats();
    if (stats.neverAsk) return;
    if (stats.insertCount < 5) return;

    const now = Date.now();
    const fourteenDays = 14 * 24 * 60 * 60 * 1000;
    if (stats.lastPromptedAt && (now - stats.lastPromptedAt) < fourteenDays) return;

    const { npsWidget, npsStarRating } = state.elements;
    if (!npsWidget.classList.contains('hidden')) return; // already showing, don't re-attach listeners
    npsWidget.classList.remove('hidden');

    npsStarRating.querySelectorAll('.star').forEach(star => {
        star.addEventListener('click', () => {
            const rating = parseInt(star.dataset.val);
            npsStarRating.querySelectorAll('.star').forEach((s, i) => {
                s.classList.toggle('active', i < rating);
            });
            sendFeedback({ type: 'nps', rating });
            saveFeedbackStats({ ...getFeedbackStats(), lastPromptedAt: now });
            npsWidget.innerHTML = '<span class="feedback-thanks" style="display:block;text-align:center;padding:8px 0;">Thanks for rating! 🙏</span>';
            setTimeout(() => npsWidget.classList.add('hidden'), 2500);
        }, { once: true });
    });

    document.getElementById('npsDismissBtn').addEventListener('click', () => {
        npsWidget.classList.add('hidden');
    }, { once: true });

    document.getElementById('npsNeverBtn').addEventListener('click', () => {
        saveFeedbackStats({ ...getFeedbackStats(), neverAsk: true });
        npsWidget.classList.add('hidden');
    }, { once: true });
}

// ============================================
// UTILITY
// ============================================

function setProcessing(isProcessing) {
    state.isProcessing = isProcessing;
    updateInputState();
}

// A book that's still uploading/processing would silently be left out of the request
// (only 'ready' books send book-ids), so sending waits for it. Typing stays enabled.
function isBookPending() {
    const status = state.attachedBook?.status;
    return status === 'uploading' || status === 'processing';
}

// Single source for the input's state — both a running generation and a pending book affect it,
// so each of setProcessing and renderBookChip calling this keeps them from overriding each other.
function updateInputState() {
    const { messageInput, bookChipCancelBtn } = state.elements;
    if (!messageInput) return;

    // A running request may be using the ready book — its answer would carry the book into the
    // history after it was removed, so removal waits (an uploading/processing book can still be cancelled)
    if (bookChipCancelBtn) {
        bookChipCancelBtn.disabled = state.isProcessing && state.attachedBook?.status === 'ready';
    }

    messageInput.disabled = state.isProcessing;
    if (state.isProcessing) {
        messageInput.placeholder = 'Generating content...';
    } else if (isBookPending()) {
        messageInput.placeholder = `Waiting for "${state.attachedBook.filename}" to finish processing…`;
    } else {
        messageInput.placeholder = 'Type your request...';
    }
}

// Feedback when a send is blocked by a pending book — the placeholder isn't visible once there's text
function flagBookPending() {
    const { bookChip } = state.elements;
    if (!bookChip) return;
    bookChip.classList.remove('book-chip-blocked');
    void bookChip.offsetWidth; // restart the animation on repeated presses
    bookChip.classList.add('book-chip-blocked');
}

function appendToChatBody(element) {
    const { chatBody, welcomeState } = state.elements;
    welcomeState.classList.add('hidden');
    chatBody.appendChild(element);
    chatBody.scrollTop = chatBody.scrollHeight;
}

function insertBeforePreview(element) {
    const { chatBody, welcomeState } = state.elements;
    welcomeState.classList.add('hidden');

    if (state.previewElement && state.previewElement.parentNode === chatBody) {
        chatBody.insertBefore(element, state.previewElement);
    } else {
        chatBody.appendChild(element);
    }

    if (state.previewElement) {
        state.previewElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}
