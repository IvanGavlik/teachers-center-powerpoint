/* global Office */

/**
 * Teachers Center Dialog - Component Logic and State Management
 * Implements Claude Code-style preview-then-execute workflow
 */

// ============================================
// STATE MANAGEMENT
// ============================================

const state = {
    // Chat state
    messages: [],
    isProcessing: false,
    pendingRequest: null,

    // Preview state
    slides: [],
    currentSlideIndex: 0,
    skippedSlides: new Set(),
    isInPreviewMode: false,

    // Progress state
    progressElement: null,

    // Deduplication - track last processed message
    lastMessageId: null,

    // UI elements (cached after init)
    elements: {}
};

// ============================================
// INITIALIZATION
// ============================================

Office.onReady((info) => {
    console.log('Dialog: Office.onReady called', info);
    initializeDialog();
});

function initializeDialog() {
    // Cache DOM elements
    state.elements = {
        chatBody: document.getElementById('chatBody'),
        welcomeState: document.getElementById('welcomeState'),
        previewArea: document.getElementById('previewArea'),
        messageInput: document.getElementById('messageInput'),
        sendBtn: document.getElementById('sendBtn'),
        quickActions: document.getElementById('quickActions'),
        closeBtn: document.getElementById('closeBtn'),
        settingsBtn: document.getElementById('settingsBtn'),
        newChatBtn: document.getElementById('newChatBtn'),
        menuBtn: document.getElementById('menuBtn')
    };

    // Set up event listeners
    setupEventListeners();

    // Listen for messages from parent (commands.js)
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        handleParentMessage
    );

    console.log('Dialog initialized');
}

function setupEventListeners() {
    const { messageInput, sendBtn, quickActions, closeBtn, settingsBtn, newChatBtn, menuBtn } = state.elements;

    // Send button
    sendBtn.addEventListener('click', handleSend);

    // Enter key to send (plain Enter sends, Shift+Enter for newline)
    messageInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    });

    // Auto-resize textarea
    messageInput.addEventListener('input', () => {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.min(messageInput.scrollHeight, 120) + 'px';
    });

    // Quick actions - fill input only
    quickActions.addEventListener('click', (e) => {
        const btn = e.target.closest('.quick-action');
        if (btn) {
            const prompt = btn.dataset.prompt || '';
            messageInput.value = prompt;
            messageInput.focus();
            // Place cursor at end
            messageInput.setSelectionRange(prompt.length, prompt.length);
            // Trigger resize
            messageInput.dispatchEvent(new Event('input'));
        }
    });

    // Header buttons
    closeBtn.addEventListener('click', handleClose);
    newChatBtn.addEventListener('click', handleNewChat);
    settingsBtn.addEventListener('click', () => console.log('Settings clicked'));
    menuBtn.addEventListener('click', () => console.log('Menu clicked'));

    // Global keyboard shortcuts
    document.addEventListener('keydown', handleGlobalKeydown);
}

// ============================================
// MESSAGE HANDLING
// ============================================

function handleSend() {
    const { messageInput } = state.elements;
    const content = messageInput.value.trim();

    if (!content || state.isProcessing) return;

    // Reset deduplication for new request
    state.lastMessageId = null;

    // Store the pending request (don't show in chat yet)
    state.pendingRequest = content;

    // Hide welcome state
    state.elements.welcomeState.classList.add('hidden');

    // Send to parent for processing
    sendToParent({
        type: 'generate',
        content: content,
        category: detectCategory(content)
    });

    // Clear input
    messageInput.value = '';
    messageInput.style.height = 'auto';

    // Show progress in preview area
    state.isProcessing = true;
    showProgressInPreviewArea('Generating content...', 0);
}

function handleClose() {
    sendToParent({ type: 'close' });
}

function handleNewChat() {
    // Clear all messages
    state.messages = [];
    state.pendingRequest = null;
    state.currentSlideIndex = 0;
    state.skippedSlides.clear();
    state.isProcessing = false;

    // Hide preview area
    hidePreviewArea();

    // Clear chat body and show welcome
    const { chatBody, welcomeState } = state.elements;
    chatBody.innerHTML = '';
    chatBody.appendChild(welcomeState);
    welcomeState.classList.remove('hidden');
}

function detectCategory(content) {
    const lower = content.toLowerCase();
    if (lower.includes('vocab') || lower.includes('word')) return 'vocabulary';
    if (lower.includes('quiz') || lower.includes('test')) return 'quiz';
    if (lower.includes('grammar')) return 'grammar';
    if (lower.includes('homework') || lower.includes('exercise')) return 'homework';
    return 'general';
}

// ============================================
// PARENT COMMUNICATION
// ============================================

function sendToParent(message) {
    try {
        Office.context.ui.messageParent(JSON.stringify(message));
    } catch (error) {
        console.error('Failed to send message to parent:', error);
    }
}

function handleParentMessage(arg) {
    console.log('Message from parent:', arg.message);

    try {
        const message = JSON.parse(arg.message);

        // Deduplicate preview messages (Office.js can deliver messages twice)
        if (message.type === 'preview') {
            const messageId = JSON.stringify(message.slides?.map(s => s.title));
            if (state.lastMessageId === messageId) {
                console.log('Duplicate preview message ignored');
                return;
            }
            state.lastMessageId = messageId;
        }

        switch (message.type) {
            case 'progress':
                updateProgressInPreviewArea(message.stage, message.percent);
                break;

            case 'preview':
                showSlidePreview(message.slides, message.summary);
                break;

            case 'insertProgress':
                updateProgressInPreviewArea(
                    `Inserting slide ${message.current} of ${message.total}...`,
                    (message.current / message.total) * 100
                );
                break;

            case 'success':
                hideProgress();
                showSuccess(message.message);
                state.isProcessing = false;
                break;

            case 'error':
                hideProgress();
                showError(message.message);
                state.isProcessing = false;
                break;

            default:
                console.log('Unknown message type:', message.type);
        }
    } catch (error) {
        console.error('Failed to parse parent message:', error);
    }
}

// ============================================
// UI COMPONENTS - Messages
// ============================================

function addUserMessage(content) {
    const template = document.getElementById('userMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-user');
    messageEl.textContent = content;

    appendToChatBody(messageEl);
    state.messages.push({ type: 'user', content });
}

function addAIMessage(content) {
    const template = document.getElementById('aiMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-ai');
    messageEl.textContent = content;

    appendToChatBody(messageEl);
    state.messages.push({ type: 'ai', content });
}

// ============================================
// UI COMPONENTS - Progress
// ============================================

function showProgress(status, percent) {
    const template = document.getElementById('progressMessageTemplate');
    const clone = template.content.cloneNode(true);
    const progressEl = clone.querySelector('.message-progress');

    progressEl.querySelector('.progress-status').textContent = status;
    progressEl.querySelector('.progress-bar').style.width = `${percent}%`;
    progressEl.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;

    // Store reference for updates
    state.progressElement = progressEl;
    appendToChatBody(progressEl);
}

function updateProgress(status, percent) {
    if (!state.progressElement) {
        showProgress(status, percent);
        return;
    }

    state.progressElement.querySelector('.progress-status').textContent = status;
    state.progressElement.querySelector('.progress-bar').style.width = `${percent}%`;
    state.progressElement.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;
}

function hideProgress() {
    if (state.progressElement && state.progressElement.parentNode) {
        state.progressElement.remove();
    }
    state.progressElement = null;
}

function showProgressInPreviewArea(status, percent) {
    const { previewArea } = state.elements;

    // Clear and show preview area
    previewArea.innerHTML = '';
    previewArea.classList.remove('hidden');

    const template = document.getElementById('progressMessageTemplate');
    const clone = template.content.cloneNode(true);
    const progressEl = clone.querySelector('.message-progress');

    progressEl.querySelector('.progress-status').textContent = status;
    progressEl.querySelector('.progress-bar').style.width = `${percent}%`;
    progressEl.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;

    // Store reference for updates
    state.progressElement = progressEl;
    previewArea.appendChild(progressEl);
}

function updateProgressInPreviewArea(status, percent) {
    if (!state.progressElement) {
        showProgressInPreviewArea(status, percent);
        return;
    }

    state.progressElement.querySelector('.progress-status').textContent = status;
    state.progressElement.querySelector('.progress-bar').style.width = `${percent}%`;
    state.progressElement.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;
}

function hidePreviewArea() {
    const { previewArea } = state.elements;
    previewArea.innerHTML = '';
    previewArea.classList.add('hidden');

    // Clear preview state
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
}

// ============================================
// UI COMPONENTS - Slide Preview
// ============================================

function showSlidePreview(slides, summary) {
    const { previewArea } = state.elements;

    // Clear the preview area
    previewArea.innerHTML = '';

    state.slides = slides;
    state.currentSlideIndex = 0;
    state.skippedSlides.clear();
    state.isInPreviewMode = true;

    const template = document.getElementById('previewContainerTemplate');
    const clone = template.content.cloneNode(true);
    const previewEl = clone.querySelector('.preview-container');

    // Set summary info
    previewEl.querySelector('.preview-title-text').textContent = summary || 'Preview';
    previewEl.querySelector('.preview-count').textContent = `${slides.length} slides`;

    // Add to the dedicated preview area at the top
    previewArea.appendChild(previewEl);
    previewArea.classList.remove('hidden');

    // Store reference for navigation
    state.previewElement = previewEl;

    // Update display for first slide
    updateSlideDisplay();

    // Set up navigation button handlers
    setupPreviewNavigation(previewEl);
}

function updateSlideDisplay() {
    if (!state.previewElement || !state.slides.length) return;

    const slide = state.slides[state.currentSlideIndex];
    const previewEl = state.previewElement;

    // Update counter
    previewEl.querySelector('.slide-counter-text').textContent =
        `Slide ${state.currentSlideIndex + 1} of ${state.slides.length}`;

    // Update slide type badge
    const typeBadge = previewEl.querySelector('.slide-type-badge');
    typeBadge.textContent = slide.type || 'Content';

    // Update card content
    previewEl.querySelector('.slide-card-title').textContent = slide.title || '';
    previewEl.querySelector('.slide-card-subtitle').textContent = slide.subtitle || '';
    previewEl.querySelector('.slide-card-content').textContent = slide.content || '';

    // Handle example sentence
    const exampleEl = previewEl.querySelector('.slide-card-example');
    if (slide.example) {
        exampleEl.textContent = slide.example;
        exampleEl.classList.remove('hidden');
    } else {
        exampleEl.classList.add('hidden');
    }

    // Update navigation buttons
    const backBtn = previewEl.querySelector('#navBackBtn');
    const nextBtn = previewEl.querySelector('#navNextBtn');

    backBtn.disabled = state.currentSlideIndex === 0;

    // Change "Next" to "Done" on last slide
    if (state.currentSlideIndex === state.slides.length - 1) {
        nextBtn.innerHTML = `
            Done
            <span class="material-icons" style="font-size: 18px;">check</span>
            <span class="nav-shortcut">Enter</span>
        `;
    } else {
        nextBtn.innerHTML = `
            Next
            <span class="material-icons" style="font-size: 18px;">arrow_forward</span>
            <span class="nav-shortcut">Enter</span>
        `;
    }

    // Update Insert All button text
    const nonSkipped = state.slides.length - state.skippedSlides.size;
    previewEl.querySelector('.insert-btn-text').textContent =
        `Insert ${nonSkipped} Slide${nonSkipped !== 1 ? 's' : ''}`;
}

function setupPreviewNavigation(previewEl) {
    previewEl.querySelector('#navBackBtn').addEventListener('click', navigateBack);
    previewEl.querySelector('#navSkipBtn').addEventListener('click', skipSlide);
    previewEl.querySelector('#navEditBtn').addEventListener('click', editSlide);
    previewEl.querySelector('#navNextBtn').addEventListener('click', navigateNext);
    previewEl.querySelector('#insertAllBtn').addEventListener('click', insertAllSlides);
}

function navigateBack() {
    if (state.currentSlideIndex > 0) {
        state.currentSlideIndex--;
        updateSlideDisplay();
    }
}

function navigateNext() {
    if (state.currentSlideIndex < state.slides.length - 1) {
        state.currentSlideIndex++;
        updateSlideDisplay();
    } else {
        // On last slide, "Done" goes to insert confirmation
        showInsertConfirmation();
    }
}

function skipSlide() {
    state.skippedSlides.add(state.currentSlideIndex);
    if (state.currentSlideIndex < state.slides.length - 1) {
        navigateNext();
    } else {
        updateSlideDisplay();
    }
}

function editSlide() {
    // Focus input with current slide info for editing
    const slide = state.slides[state.currentSlideIndex];
    const { messageInput } = state.elements;

    messageInput.value = `Edit slide ${state.currentSlideIndex + 1}: ${slide.title || ''}`;
    messageInput.focus();
    messageInput.dispatchEvent(new Event('input'));
}

function showInsertConfirmation() {
    const nonSkipped = state.slides.length - state.skippedSlides.size;

    if (nonSkipped === 0) {
        showError('All slides have been skipped. Nothing to insert.');
        return;
    }

    // Change button state to show ready
    if (state.previewElement) {
        const insertBtn = state.previewElement.querySelector('#insertAllBtn');
        insertBtn.focus();
    }
}

function insertAllSlides() {
    // Get slides to insert (excluding skipped)
    const slidesToInsert = state.slides.filter((_, index) => !state.skippedSlides.has(index));

    if (slidesToInsert.length === 0) {
        showError('No slides selected for insertion.');
        return;
    }

    // Show progress in preview area (replaces slide preview)
    state.isProcessing = true;
    showProgressInPreviewArea('Starting slide insertion...', 0);

    // Send to parent for insertion
    sendToParent({
        type: 'insert',
        slides: slidesToInsert
    });

    // Exit preview mode
    state.isInPreviewMode = false;
}

// ============================================
// UI COMPONENTS - Success/Error
// ============================================

function showSuccess(message) {
    // Hide and clear preview area first
    hidePreviewArea();

    // Add the original request to chat history
    if (state.pendingRequest) {
        addUserMessage(state.pendingRequest);
        state.pendingRequest = null;
    }

    // Add success message
    const template = document.getElementById('successMessageTemplate');
    const clone = template.content.cloneNode(true);
    const successEl = clone.querySelector('.message-success');
    successEl.querySelector('.success-text').textContent = message;

    appendToChatBody(successEl);
}

function showError(message) {
    // Hide and clear preview area first
    hidePreviewArea();

    // Add the original request to chat history
    if (state.pendingRequest) {
        addUserMessage(state.pendingRequest);
        state.pendingRequest = null;
    }

    // Add error message
    const template = document.getElementById('errorMessageTemplate');
    const clone = template.content.cloneNode(true);
    const errorEl = clone.querySelector('.message-error');
    errorEl.querySelector('.error-text').textContent = message;

    appendToChatBody(errorEl);
}

// ============================================
// KEYBOARD NAVIGATION
// ============================================

function handleGlobalKeydown(e) {
    // Only handle shortcuts in preview mode
    if (!state.isInPreviewMode || state.isProcessing) return;

    // Don't capture when typing in input
    if (document.activeElement === state.elements.messageInput) return;

    switch (e.key.toLowerCase()) {
        case 'enter':
            e.preventDefault();
            navigateNext();
            break;

        case 's':
            e.preventDefault();
            skipSlide();
            break;

        case 'e':
            e.preventDefault();
            editSlide();
            break;

        case 'a':
            e.preventDefault();
            insertAllSlides();
            break;

        case 'escape':
            e.preventDefault();
            // Focus input to allow typing refinements
            state.elements.messageInput.focus();
            break;

        case 'arrowleft':
        case 'backspace':
            if (state.currentSlideIndex > 0) {
                e.preventDefault();
                navigateBack();
            }
            break;

        case 'arrowright':
            if (state.currentSlideIndex < state.slides.length - 1) {
                e.preventDefault();
                navigateNext();
            }
            break;
    }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function appendToChatBody(element) {
    const { chatBody, welcomeState } = state.elements;

    // Make sure welcome state is hidden when adding messages
    welcomeState.classList.add('hidden');

    chatBody.appendChild(element);

    // Scroll to bottom
    chatBody.scrollTop = chatBody.scrollHeight;
}

function scrollToBottom() {
    const { chatBody } = state.elements;
    chatBody.scrollTop = chatBody.scrollHeight;
}

// ============================================
// DEMO/TESTING - Generate sample slides
// (Remove in production - backend will provide real data)
// ============================================

function generateDemoSlides(content) {
    // Parse simple vocabulary request
    const slides = [
        {
            type: 'Title',
            title: 'German Food Vocabulary',
            subtitle: 'Level: A1 • 5 words',
            content: ''
        },
        {
            type: 'Vocabulary',
            title: 'der Apfel',
            subtitle: 'the apple',
            content: 'A common fruit enjoyed in Germany.',
            example: '"Ich esse einen Apfel." (I eat an apple.)'
        },
        {
            type: 'Vocabulary',
            title: 'das Brot',
            subtitle: 'the bread',
            content: 'Germans consume more bread varieties than any other country.',
            example: '"Möchtest du Brot?" (Would you like bread?)'
        },
        {
            type: 'Vocabulary',
            title: 'die Milch',
            subtitle: 'the milk',
            content: 'Used in many German breakfast items.',
            example: '"Die Milch ist frisch." (The milk is fresh.)'
        },
        {
            type: 'Vocabulary',
            title: 'der Käse',
            subtitle: 'the cheese',
            content: 'Germany produces over 450 types of cheese.',
            example: '"Ich mag Käse sehr." (I like cheese a lot.)'
        },
        {
            type: 'Vocabulary',
            title: 'das Wasser',
            subtitle: 'the water',
            content: 'Germans often drink sparkling water (Sprudel).',
            example: '"Ein Glas Wasser, bitte." (A glass of water, please.)'
        }
    ];

    return slides;
}
