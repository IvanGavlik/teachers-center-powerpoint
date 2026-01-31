/* global Office */

/**
 * Teachers Center Dialog - Component Logic and State Management
 * Implements Claude Code-style preview-then-execute workflow
 */

// Import CSS for webpack bundling
import './dialog.css';

// ============================================
// WEBSOCKET CONFIGURATION
// ============================================

// URL injected by webpack based on environment (dev/prod)
const WS_URL = process.env.WS_URL;
const USER_ID = 'user-123';  // TODO: implement proper user management
const CHANNEL_NAME = 'powerpoint-dialog';
const MAX_RECONNECT_ATTEMPTS = 3;

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

    // Settings
    settings: {
        language: 'English',
        level: 'B1',
        className: '',
        nativeLanguage: '',
        ageGroup: ''
    },

    // WebSocket state
    ws: null,
    wsState: 'disconnected', // 'disconnected' | 'connecting' | 'connected'
    reconnectAttempts: 0,
    conversationId: null,

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
        quickActions: document.getElementById('quickActions'),
        closeBtn: document.getElementById('closeBtn'),
        // TODO: implement in version 2
        // settingsBtn: document.getElementById('settingsBtn'),
        newChatBtn: document.getElementById('newChatBtn'),
        // TODO: implement in version 2
        // contextBadge: document.getElementById('contextBadge'),
        // Settings modal elements
        settingsModal: document.getElementById('settingsModal'),
        closeModalBtn: document.getElementById('closeModalBtn'),
        cancelSettingsBtn: document.getElementById('cancelSettingsBtn'),
        saveSettingsBtn: document.getElementById('saveSettingsBtn'),
        settingsLanguage: document.getElementById('settingsLanguage'),
        settingsLevel: document.getElementById('settingsLevel'),
        settingsClassName: document.getElementById('settingsClassName'),
        settingsNativeLanguage: document.getElementById('settingsNativeLanguage'),
        settingsAgeGroup: document.getElementById('settingsAgeGroup')
    };

    // Load saved settings
    loadSettings();

    // Set up event listeners
    setupEventListeners();

    // Listen for messages from parent (commands.js)
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        handleParentMessage
    );

    // Connect to WebSocket backend (delayed to allow page to fully load)
    setTimeout(() => {
        connectWebSocket();
    }, 500);

    console.log('Dialog initialized');
}

function setupEventListeners() {
    const { messageInput, quickActions, closeBtn, newChatBtn } = state.elements;

    // Handle window resize - keep preview visible
    window.addEventListener('resize', () => {
        if (state.isInPreviewMode && state.previewElement) {
            state.previewElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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
    // TODO: implement in version 2
    // settingsBtn.addEventListener('click', openSettingsModal);

    // TODO: implement in version 2
    // Context badge also opens settings
    // const { contextBadge } = state.elements;
    // if (contextBadge) {
    //     contextBadge.addEventListener('click', openSettingsModal);
    // }

    // Settings modal buttons
    const { closeModalBtn, cancelSettingsBtn, saveSettingsBtn, settingsModal } = state.elements;
    closeModalBtn.addEventListener('click', closeSettingsModal);
    cancelSettingsBtn.addEventListener('click', closeSettingsModal);
    saveSettingsBtn.addEventListener('click', saveSettings);

    // Close modal when clicking overlay
    settingsModal.addEventListener('click', (e) => {
        if (e.target === settingsModal) {
            closeSettingsModal();
        }
    });

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

    // If there's an existing preview, dismiss it first
    if (state.isInPreviewMode && state.slides.length > 0) {
        const skippedCount = state.slides.length - state.skippedSlides.size;
        dismissPreview(`${skippedCount} slide${skippedCount !== 1 ? 's' : ''} not inserted`);
    }

    // Reset deduplication for new request
    state.lastMessageId = null;

    // Hide welcome state
    state.elements.welcomeState.classList.add('hidden');

    // Show user message immediately
    addUserMessage(content);

    // Clear input
    messageInput.value = '';
    messageInput.style.height = 'auto';

    // Show progress in preview area
    state.isProcessing = true;
    showProgressInPreviewArea('Generating content...', 0);

    // Send to WebSocket backend
    const sent = sendWebSocketMessage({
        type: 'generate',
        content: content,
        category: detectCategory(content)
    });

    // If WebSocket send failed, try to reconnect and show error
    if (!sent) {
        connectWebSocket();
    }
}

function handleClose() {
    disconnectWebSocket();
    sendToParent({ type: 'close' });
}

function handleNewChat() {
    // Clear all messages
    state.messages = [];
    state.pendingRequest = null;
    state.currentSlideIndex = 0;
    state.skippedSlides.clear();
    state.isProcessing = false;
    state.conversationId = null; // New conversation ID will be generated on next message

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
// WEBSOCKET CONNECTION
// ============================================

function connectWebSocket() {
    if (state.ws && state.wsState === 'connected') {
        console.log('WebSocket already connected');
        return;
    }

    // Check if WebSocket URL is configured
    if (!WS_URL) {
        console.warn('WebSocket URL not configured, skipping connection');
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
        console.log('WebSocket message received:', event.data);
        try {
            handleWebSocketMessage(event.data);
        } catch (error) {
            console.error('Error handling WebSocket message:', error);
        }
    };

    state.ws.onclose = (event) => {
        console.log('WebSocket closed:', event.code, event.reason);
        state.wsState = 'disconnected';
        state.ws = null;

        // Attempt reconnection if not intentional close
        if (event.code !== 1000 && state.reconnectAttempts < MAX_RECONNECT_ATTEMPTS) {
            state.reconnectAttempts++;
            console.log(`Reconnecting (attempt ${state.reconnectAttempts}/${MAX_RECONNECT_ATTEMPTS})...`);
            setTimeout(connectWebSocket, 2000 * state.reconnectAttempts);
        }
    };

    state.ws.onerror = (error) => {
        console.warn('WebSocket error (backend may not be running):', error);
        // Don't set state here - onclose will be called next
    };
}

function disconnectWebSocket() {
    if (state.ws) {
        state.ws.close(1000, 'User closed dialog');
        state.ws = null;
        state.wsState = 'disconnected';
    }
}

function sendWebSocketMessage(message) {
    if (!state.ws || state.wsState !== 'connected') {
        console.error('WebSocket not connected, cannot send message. State:', state.wsState);
        showError('Not connected to server. Make sure the backend is running and refresh the dialog.');
        state.isProcessing = false;
        return false;
    }

    try {
        // Generate conversation ID if not exists
        if (!state.conversationId) {
            state.conversationId = `conv-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
        }

        // Format message for backend
        const wsMessage = {
            'user-id': USER_ID,
            'channel-name': CHANNEL_NAME,
            'conversation-id': state.conversationId,
            type: message.type || 'generate',
            content: message.content,
            requirements: {
                category: message.category || 'general',
                language: state.settings.language,
                level: state.settings.level,
                nativeLanguage: state.settings.nativeLanguage || null,
                ageGroup: state.settings.ageGroup || null,
                className: state.settings.className || null
            }
        };

        console.log('Sending WebSocket message:', wsMessage);
        state.ws.send(JSON.stringify(wsMessage));
        return true;
    } catch (error) {
        console.error('Failed to send WebSocket message:', error);
        showError('Failed to send message. Please try again.');
        state.isProcessing = false;
        return false;
    }
}

function handleWebSocketMessage(data) {
    try {
        const message = JSON.parse(data);
        console.log('Parsed WebSocket message:', message);

        // Handle different message types from backend
        // Check for known fields first (backend may not always send 'type')
        if (message['requirements-not-meet']) {
            // Backend needs more information
            hideProgress();
            addAIMessage(message['requirements-not-meet']);
            state.isProcessing = false;
            return;
        }

        if (message.slides || message.data) {
            // Backend sent slides
            const slides = transformBackendSlides(message.slides || message.data);
            if (slides && slides.length > 0) {
                showSlidePreview(slides, message.summary || 'Generated Slides');
            } else {
                showError('No slides generated. Please try a different request.');
                state.isProcessing = false;
            }
            return;
        }

        if (message.error || message.message) {
            // Backend sent an error or info message
            hideProgress();
            if (message.error) {
                showError(message.error);
            } else {
                addAIMessage(message.message);
            }
            state.isProcessing = false;
            return;
        }

        // Handle by explicit type field
        switch (message.type) {
            case 'progress':
                updateProgressInPreviewArea(message.stage || message.message, message.percent || 0);
                break;

            case 'slides':
            case 'result':
                const slides = transformBackendSlides(message.slides || message.data);
                if (slides && slides.length > 0) {
                    showSlidePreview(slides, message.summary || 'Generated Slides');
                } else {
                    showError('No slides generated. Please try a different request.');
                    state.isProcessing = false;
                }
                break;

            case 'error':
                hideProgress();
                showError(message.message || 'An error occurred');
                state.isProcessing = false;
                break;

            case 'connected':
                console.log('Backend confirmed connection');
                break;

            default:
                console.log('Unknown WebSocket message type:', message.type);
                // If we got here with no handler, stop processing state
                hideProgress();
                state.isProcessing = false;
        }
    } catch (error) {
        console.error('Failed to parse WebSocket message:', error);
        hideProgress();
        state.isProcessing = false;
    }
}

function transformBackendSlides(backendSlides) {
    // Transform backend slide format to our internal format
    // Backend format may vary - adjust as needed
    if (!backendSlides || !Array.isArray(backendSlides)) {
        return [];
    }

    return backendSlides.map(slide => ({
        type: slide.type || 'Content',
        title: slide.title || '',
        subtitle: slide.subtitle || '',
        content: slide.content || slide.body || '',
        example: slide.example || slide['example-sentence'] || ''
    }));
}

// ============================================
// PARENT COMMUNICATION (for PowerPoint API)
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
                console.log('Updating progress:', message.stage, message.percent);
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
    const template = document.getElementById('progressMessageTemplate');
    const clone = template.content.cloneNode(true);
    const progressEl = clone.querySelector('.message-progress');

    progressEl.querySelector('.progress-status').textContent = status;
    progressEl.querySelector('.progress-bar').style.width = `${percent}%`;
    progressEl.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;

    // Store reference for updates
    state.progressElement = progressEl;

    // Append to chat body (after user message)
    appendToChatBody(progressEl);
}

function updateProgressInPreviewArea(status, percent) {
    console.log('updateProgressInPreviewArea called:', status, percent, 'progressElement exists:', !!state.progressElement);
    if (!state.progressElement) {
        console.log('No progress element, creating new one');
        showProgressInPreviewArea(status, percent);
        return;
    }

    state.progressElement.querySelector('.progress-status').textContent = status;
    state.progressElement.querySelector('.progress-bar').style.width = `${percent}%`;
    state.progressElement.querySelector('.progress-text').textContent = `${Math.round(percent)}%`;
}

function hidePreviewArea() {
    // Remove preview element from chat body if exists
    if (state.previewElement && state.previewElement.parentNode) {
        state.previewElement.remove();
    }

    // Clear preview state
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
}

function dismissPreview(message) {
    // Remove preview element
    if (state.previewElement && state.previewElement.parentNode) {
        state.previewElement.remove();
    }

    // Show dismiss message
    const dismissEl = document.createElement('div');
    dismissEl.className = 'message-dismissed fade-in';
    dismissEl.innerHTML = `
        <span class="material-icons">info</span>
        <span>${message}</span>
    `;
    appendToChatBody(dismissEl);

    // Clear preview state
    state.slides = [];
    state.previewElement = null;
    state.skippedSlides.clear();
    state.isInPreviewMode = false;
}

// ============================================
// UI COMPONENTS - Slide Preview
// ============================================

function showSlidePreview(slides, summary) {
    console.log('showSlidePreview called with', slides.length, 'slides');

    // Remove progress element if exists
    if (state.progressElement && state.progressElement.parentNode) {
        state.progressElement.remove();
        state.progressElement = null;
    }

    state.slides = slides;
    state.currentSlideIndex = 0;
    state.skippedSlides.clear();
    state.isInPreviewMode = true;
    state.isProcessing = false;  // Ready for navigation
    console.log('isInPreviewMode set to TRUE, isProcessing set to FALSE');

    const template = document.getElementById('previewContainerTemplate');
    const clone = template.content.cloneNode(true);
    const previewEl = clone.querySelector('.preview-container');

    // Set summary info
    previewEl.querySelector('.preview-title-text').textContent = summary || 'Preview';
    previewEl.querySelector('.preview-count').textContent = `${slides.length} slides`;

    // Store reference for navigation
    state.previewElement = previewEl;

    // Append to chat body (after user message)
    appendToChatBody(previewEl);

    // Update display for first slide
    updateSlideDisplay();

    // Set up navigation button handlers
    setupPreviewNavigation(previewEl);

    // Scroll to show preview
    setTimeout(() => {
        previewEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 100);
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

    // Change "Next" to "Insert X Slides" on last slide
    const nonSkipped = state.slides.length - state.skippedSlides.size;
    if (state.currentSlideIndex === state.slides.length - 1) {
        nextBtn.innerHTML = `
            Insert ${nonSkipped} Slide${nonSkipped !== 1 ? 's' : ''}
            <span class="material-icons" style="font-size: 18px;">playlist_add</span>
            <span class="nav-shortcut">Enter</span>
        `;
    } else {
        nextBtn.innerHTML = `
            Next
            <span class="material-icons" style="font-size: 18px;">arrow_forward</span>
            <span class="nav-shortcut">Enter</span>
        `;
    }
}

function setupPreviewNavigation(previewEl) {
    previewEl.querySelector('#navBackBtn').addEventListener('click', navigateBack);
    previewEl.querySelector('#navSkipBtn').addEventListener('click', skipSlide);
    previewEl.querySelector('#navEditBtn').addEventListener('click', editSlide);
    previewEl.querySelector('#navNextBtn').addEventListener('click', navigateNext);
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
        // On last slide, "Done" inserts all slides
        insertAllSlides();
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

    // Hide preview immediately
    hidePreviewArea();

    // Show progress
    state.isProcessing = true;
    showProgressInPreviewArea('Inserting slides...', 0);

    // Send to parent for insertion
    sendToParent({
        type: 'insert',
        slides: slidesToInsert
    });
}

// ============================================
// UI COMPONENTS - Success/Error
// ============================================

function showSuccess(message) {
    // Hide and clear preview area first
    hidePreviewArea();

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
    // Handle Enter key globally
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();

        const hasText = state.elements.messageInput.value.trim().length > 0;

        if (hasText && !state.isProcessing) {
            // Has text in input → send message
            handleSend();
        } else if (state.isInPreviewMode && !state.isProcessing) {
            // Empty input + in preview mode → Next
            navigateNext();
        }
        return;
    }

    // Other shortcuts only work in preview mode and when NOT typing in input
    if (!state.isInPreviewMode || state.isProcessing) return;
    if (document.activeElement === state.elements.messageInput) return;

    switch (e.key.toLowerCase()) {
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
// SETTINGS FUNCTIONS
// ============================================

function openSettingsModal() {
    const { settingsModal } = state.elements;
    settingsModal.classList.remove('hidden');
}

function closeSettingsModal() {
    const { settingsModal } = state.elements;
    settingsModal.classList.add('hidden');
    // Reset form to saved values
    loadSettingsToForm();
}

function loadSettings() {
    // Load settings from localStorage
    const savedSettings = localStorage.getItem('teachersCenterSettings');
    if (savedSettings) {
        state.settings = JSON.parse(savedSettings);
    } else {
        // Default settings
        state.settings = {
            language: 'English',
            level: 'B1',
            className: '',
            nativeLanguage: '',
            ageGroup: ''
        };
    }

    // Apply settings to form and context badge
    loadSettingsToForm();
    updateContextBadge();
}

function loadSettingsToForm() {
    const { settingsLanguage, settingsLevel, settingsClassName, settingsNativeLanguage, settingsAgeGroup } = state.elements;

    settingsLanguage.value = state.settings.language;
    settingsLevel.value = state.settings.level;
    settingsClassName.value = state.settings.className;
    settingsNativeLanguage.value = state.settings.nativeLanguage;
    settingsAgeGroup.value = state.settings.ageGroup;
}

function saveSettings() {
    const { settingsLanguage, settingsLevel, settingsClassName, settingsNativeLanguage, settingsAgeGroup } = state.elements;

    // Update state
    state.settings = {
        language: settingsLanguage.value,
        level: settingsLevel.value,
        className: settingsClassName.value,
        nativeLanguage: settingsNativeLanguage.value,
        ageGroup: settingsAgeGroup.value
    };

    // Save to localStorage
    localStorage.setItem('teachersCenterSettings', JSON.stringify(state.settings));

    // Update context badge
    updateContextBadge();

    // Close modal
    closeSettingsModal();
}

function updateContextBadge() {
    const { contextBadge } = state.elements;
    if (contextBadge) {
        contextBadge.textContent = `${state.settings.level} ${state.settings.language}`;
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
