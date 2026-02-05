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

    // Edit mode state
    originalRequest: null,      // The user's original prompt that generated current slides
    editingSlideIndex: null,    // Which slide is being edited (null = not editing)
    isEditMode: false,          // True when user clicked Edit and is typing

    // Progress state
    progressElement: null,
    cancelled: false,
    staleResponses: 0,

    // Deduplication - track last processed message
    lastMessageId: null,

    // Type selection (required before sending) - vocabulary is default
    selectedType: 'vocabulary', // 'vocabulary' | 'grammar' | 'quiz' | 'homework'

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
        inputContainer: document.getElementById('inputContainer'),
        closeBtn: document.getElementById('closeBtn'),
        // TODO: implement in version 2
        // settingsBtn: document.getElementById('settingsBtn'),
        newChatBtn: document.getElementById('newChatBtn'),
        // TODO: implement in version 2
        // contextBadge: document.getElementById('contextBadge'),
        // Type selector elements
        typeSelector: document.getElementById('typeSelector'),
        typeOptions: document.getElementById('typeOptions'),
        typeError: document.getElementById('typeError'),
        // Edit badge elements
        editBadge: document.getElementById('editBadge'),
        editBadgeSlideNum: document.getElementById('editBadgeSlideNum'),
        editBadgeCancelBtn: document.getElementById('editBadgeCancelBtn'),
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

    // Set initial placeholder for default type (vocabulary)
    const defaultTypeBtn = state.elements.typeOptions.querySelector('.type-option.selected');
    if (defaultTypeBtn && defaultTypeBtn.dataset.placeholder) {
        state.elements.messageInput.placeholder = defaultTypeBtn.dataset.placeholder;
    }

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
    const { messageInput, closeBtn, newChatBtn, typeOptions } = state.elements;

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
        // Clear validation error when user starts typing
        clearTypeValidationError();
    });

    // Type selection - select type and optionally fill prompt
    typeOptions.addEventListener('click', (e) => {
        const btn = e.target.closest('.type-option');
        if (btn) {
            selectType(btn);
        }
    });

    // Header buttons
    closeBtn.addEventListener('click', handleClose);
    newChatBtn.addEventListener('click', handleNewChat);
    // TODO: implement in version 2
    // settingsBtn.addEventListener('click', openSettingsModal);

    // Edit badge cancel button
    const { editBadgeCancelBtn } = state.elements;
    if (editBadgeCancelBtn) {
        editBadgeCancelBtn.addEventListener('click', exitEditMode);
    }

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
// TYPE SELECTION
// ============================================

function selectType(btn) {
    const { typeOptions, messageInput } = state.elements;
    const type = btn.dataset.type;
    const placeholder = btn.dataset.placeholder || 'Type your request...';

    // Remove selected class from all options
    typeOptions.querySelectorAll('.type-option').forEach(opt => {
        opt.classList.remove('selected');
    });

    // Add selected class to clicked option
    btn.classList.add('selected');

    // Update state
    state.selectedType = type;

    // Update placeholder text
    messageInput.placeholder = placeholder;

    // Clear any validation error
    clearTypeValidationError();

    // Focus the input
    messageInput.focus();

    console.log('Type selected:', type);
}

function resetTypeSelection() {
    const { typeOptions, messageInput } = state.elements;

    // Remove selected class from all options
    typeOptions.querySelectorAll('.type-option').forEach(opt => {
        opt.classList.remove('selected');
    });

    // Reset to default (vocabulary)
    const defaultBtn = typeOptions.querySelector('[data-type="vocabulary"]');
    if (defaultBtn) {
        defaultBtn.classList.add('selected');
        state.selectedType = 'vocabulary';
        messageInput.placeholder = defaultBtn.dataset.placeholder || 'Type your request...';
    }

    // Clear any validation error
    clearTypeValidationError();
}

function showTypeValidationError() {
    const { typeError, inputContainer } = state.elements;
    typeError.classList.remove('hidden');
    inputContainer.classList.add('validation-error');
}

function clearTypeValidationError() {
    const { typeError, inputContainer } = state.elements;
    typeError.classList.add('hidden');
    inputContainer.classList.remove('validation-error');
}

function validateTypeSelection() {
    if (!state.selectedType) {
        showTypeValidationError();
        return false;
    }
    return true;
}

// ============================================
// MESSAGE HANDLING
// ============================================

function handleSend() {
    const { messageInput } = state.elements;
    const content = messageInput.value.trim();

    if (!content || state.isProcessing) return;

    // Check if we're in edit mode - route to edit handler
    if (state.isEditMode && state.editingSlideIndex !== null) {
        handleEditSend(content);
        return;
    }

    // Validate type selection (required for normal generation)
    if (!validateTypeSelection()) {
        return;
    }

    // If there's an existing preview, dismiss it first
    if (state.isInPreviewMode && state.slides.length > 0) {
        const skippedCount = state.slides.length - state.skippedSlides.size;
        dismissPreview(`${skippedCount} slide${skippedCount !== 1 ? 's' : ''} not inserted`);
    }

    // Reset state for new request
    state.lastMessageId = null;
    state.cancelled = false;

    // Hide welcome state
    state.elements.welcomeState.classList.add('hidden');

    // Show user message immediately
    addUserMessage(content);

    // Clear input
    messageInput.value = '';
    messageInput.style.height = 'auto';

    // Show progress in preview area
    setProcessing(true);
    showProgressInPreviewArea('Generating content...');

    // Save the original request for potential edit operations
    state.originalRequest = content;

    // Send to WebSocket backend with selected type
    const sent = sendWebSocketMessage({
        type: state.selectedType,
        content: content
    });

    // If WebSocket send failed, try to reconnect and show error
    if (!sent) {
        connectWebSocket();
    }
}

function handleEditSend(editInstruction) {
    // Handle sending an edit request for a specific slide
    const { messageInput } = state.elements;
    const slideIndex = state.editingSlideIndex;
    const currentSlide = state.slides[slideIndex];

    // Show user message BEFORE the preview (to keep preview visible)
    const messageContent = `Edit slide ${slideIndex + 1}: ${editInstruction}`;
    const template = document.getElementById('userMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-user');
    messageEl.textContent = messageContent;
    insertBeforePreview(messageEl);
    state.messages.push({ type: 'user', content: messageContent });

    // Stay in edit mode - user can make more edits to the same slide
    // They exit explicitly by clicking X on the edit badge

    // Clear input
    messageInput.value = '';
    messageInput.style.height = 'auto';

    // Show progress in preview area
    setProcessing(true);
    showProgressInPreviewArea('Updating slide...');

    // Send edit request to backend
    const sent = sendWebSocketMessage({
        type: 'edit',
        content: editInstruction,
        edit: {
            slideIndex: slideIndex,
            currentSlide: currentSlide,
            originalRequest: state.originalRequest,
            originalType: state.selectedType
        }
    });

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
    setProcessing(false);
    state.conversationId = null; // New conversation ID will be generated on next message

    // Reset edit mode state
    state.originalRequest = null;
    state.editingSlideIndex = null;
    state.isEditMode = false;

    // Reset type selection to default (vocabulary)
    resetTypeSelection();

    // Hide preview area
    hidePreviewArea();

    // Clear chat body and show welcome
    const { chatBody, welcomeState } = state.elements;
    chatBody.innerHTML = '';
    chatBody.appendChild(welcomeState);
    welcomeState.classList.remove('hidden');
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
        setProcessing(false);
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
            type: message.type, // Required: 'vocabulary' | 'grammar' | 'quiz' | 'homework' | 'edit'
            content: message.content,
            requirements: {
                language: state.settings.language,
                level: state.settings.level,
                nativeLanguage: state.settings.nativeLanguage || null,
                ageGroup: state.settings.ageGroup || null,
                className: state.settings.className || null
            },
            // Include edit data when present (for edit requests)
            ...(message.edit && { edit: message.edit })
        };

        console.log('Sending WebSocket message:', JSON.stringify(wsMessage, null, 2));
        state.ws.send(JSON.stringify(wsMessage));
        return true;
    } catch (error) {
        console.error('Failed to send WebSocket message:', error);
        showError('Failed to send message. Please try again.');
        setProcessing(false);
        return false;
    }
}

function handleWebSocketMessage(data) {
    try {
        const message = JSON.parse(data);
        console.log('Parsed WebSocket message:', message);

        // Ignore stale responses from cancelled requests
        if (state.staleResponses > 0) {
            if (message.type !== 'progress') {
                state.staleResponses--;
                console.log('Discarded stale response from cancelled request');
            }
            return;
        }

        // Handle different message types from backend
        // Check for known fields first (backend may not always send 'type')
        if (message['requirements-not-meet']) {
            // Backend needs more information
            hideProgress();
            addAIMessage(message['requirements-not-meet']);
            setProcessing(false);
            return;
        }

        // Handle progress updates (sent during content generation)
        if (message.type === 'progress') {
            updateProgressInPreviewArea(message.stage || message.message);
            return;
        }

        // Handle edit responses (single slide update)
        if (message.type === 'edit' && message.edit) {
            const slideIndex = message.edit.slideIndex;
            const existingSlide = state.slides[slideIndex];  // Pass existing slide to preserve title
            const transformedSlide = transformEditedSlide(message.edit.slide, state.selectedType, existingSlide);

            if (transformedSlide && slideIndex >= 0 && slideIndex < state.slides.length) {
                // Replace the edited slide in the array
                state.slides[slideIndex] = transformedSlide;
                state.currentSlideIndex = slideIndex;
                updateSlideDisplay();
                hideProgress();
                setProcessing(false);
                addAIMessage('Slide updated.');
                // Scroll to keep preview visible and focus input for next edit
                if (state.previewElement) {
                    setTimeout(() => {
                        state.previewElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                        // If still in edit mode, focus input for next edit
                        if (state.isEditMode) {
                            state.elements.messageInput.focus();
                        }
                    }, 100);
                }
            } else {
                showError('Failed to apply edit. Please try again.');
                setProcessing(false);
            }
            return;
        }

        // Route by type field from backend (vocabulary, grammar, quiz, homework)
        if (message.type) {
            const slides = transformResponseByType(message);
            if (slides && slides.length > 0) {
                const titles = {
                    vocabulary: 'Vocabulary',
                    grammar: 'Grammar',
                    quiz: 'Quiz',
                    homework: 'Homework'
                };
                showSlidePreview(slides, message.title || titles[message.type] || 'Generated Content');
            } else {
                showError(`No ${message.type} content generated. Please try a different request.`);
                setProcessing(false);
            }
            return;
        }

        // Fallback for generic slides format (legacy)
        if (message.slides || message.data) {
            const slides = transformBackendSlides(message.slides || message.data);
            if (slides && slides.length > 0) {
                showSlidePreview(slides, message.summary || 'Generated Slides');
            } else {
                showError('No slides generated. Please try a different request.');
                setProcessing(false);
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
            setProcessing(false);
            return;
        }

        // Handle by explicit type field
        switch (message.type) {
            case 'progress':
                updateProgressInPreviewArea(message.stage || message.message);
                break;

            case 'slides':
            case 'result':
                const slides = transformBackendSlides(message.slides || message.data);
                if (slides && slides.length > 0) {
                    showSlidePreview(slides, message.summary || 'Generated Slides');
                } else {
                    showError('No slides generated. Please try a different request.');
                    setProcessing(false);
                }
                break;

            case 'error':
                hideProgress();
                showError(message.message || 'An error occurred');
                setProcessing(false);
                break;

            case 'connected':
                console.log('Backend confirmed connection');
                break;

            default:
                console.log('Unknown WebSocket message type:', message.type);
                // If we got here with no handler, stop processing state
                hideProgress();
                setProcessing(false);
        }
    } catch (error) {
        console.error('Failed to parse WebSocket message:', error);
        hideProgress();
        setProcessing(false);
    }
}

function transformResponseByType(message) {
    // Route to appropriate transform function based on type
    switch (message.type) {
        case 'vocabulary':
            return transformVocabularyToSlides(message);
        case 'grammar':
            return transformGrammarToSlides(message);
        case 'quiz':
            return transformQuizToSlides(message);
        case 'homework':
            return transformHomeworkToSlides(message);
        default:
            console.warn('Unknown message type:', message.type);
            return transformBackendSlides(message.slides || message.data || []);
    }
}

function transformBackendSlides(backendSlides) {
    // Transform backend slide format to our internal format (fallback)
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
// SINGLE-ITEM TRANSFORM FUNCTIONS
// These are reused by both generation and edit flows
// ============================================

function transformSingleVocabularySlide(wordData) {
    // Transform a single vocabulary word to internal slide format
    // Input: { word, definition, translation, example }
    // Output: { type, title, subtitle, content, example }
    return {
        type: 'Vocabulary',
        title: wordData.word || wordData.title || '',
        subtitle: wordData.translation || wordData.subtitle || '',
        content: wordData.definition || wordData.content || '',
        example: wordData.example || ''
    };
}

function transformSingleGrammarSlide(slideData) {
    // Transform a single grammar slide to internal slide format
    // Input: { slide-title, content: { explanation, form, usage, examples } }
    // Output: { type, title, subtitle, content, example }
    const content = slideData.content || slideData;

    // Build content string from grammar structure
    let contentText = '';

    if (content.explanation) {
        contentText += content.explanation + '\n\n';
    }

    if (content.form) {
        contentText += '📝 Form:\n';
        if (content.form.positive) contentText += `  ✓ ${content.form.positive}\n`;
        if (content.form.negative) contentText += `  ✗ ${content.form.negative}\n`;
        if (content.form.question) contentText += `  ? ${content.form.question}\n`;
        contentText += '\n';
    }

    if (content.usage && Array.isArray(content.usage)) {
        contentText += '💡 Usage:\n';
        content.usage.forEach(u => {
            contentText += `  • ${u}\n`;
        });
    }

    // Build examples string
    let exampleText = '';
    if (content.examples && Array.isArray(content.examples)) {
        content.examples.forEach(ex => {
            if (ex.sentence) {
                exampleText += ex.sentence;
                if (ex.translation) exampleText += ` → ${ex.translation}`;
                exampleText += '\n';
            }
        });
    }

    return {
        type: 'Grammar',
        title: slideData['slide-title'] || slideData.slideTitle || slideData.title || 'Grammar Rule',
        subtitle: '',
        content: contentText.trim(),
        example: exampleText.trim()
    };
}

function transformSingleQuizSlide(questionData, title = 'Question') {
    // Transform a single quiz question to internal slide format
    // Input: { question, options, correct-answer }
    // Output: { type, title, subtitle, content, example }
    let contentText = (questionData.question || '') + '\n\n';

    if (questionData.options && Array.isArray(questionData.options)) {
        questionData.options.forEach((opt, i) => {
            const letter = String.fromCharCode(65 + i); // A, B, C, D
            contentText += `${letter}. ${opt}\n`;
        });
    }

    return {
        type: 'Quiz',
        title: questionData.title || title,
        subtitle: '',
        content: contentText.trim(),
        example: questionData['correct-answer'] ? `Answer: ${questionData['correct-answer']}` : ''
    };
}

function transformSingleHomeworkSlide(taskData, title = 'Task') {
    // Transform a single homework task to internal slide format
    // Input: { instruction, items }
    // Output: { type, title, subtitle, content, example }
    let contentText = '';

    if (taskData.instruction) {
        contentText += `📋 ${taskData.instruction}\n\n`;
    }

    if (taskData.items && Array.isArray(taskData.items)) {
        taskData.items.forEach((item, i) => {
            contentText += `${i + 1}. ${item}\n`;
        });
    }

    return {
        type: 'Homework',
        title: taskData.title || title,
        subtitle: taskData.instruction || '',
        content: contentText.trim(),
        example: ''
    };
}

// ============================================
// MULTI-SLIDE TRANSFORM FUNCTIONS
// These create title slide + content slides for generation
// ============================================

function transformVocabularyToSlides(vocabData) {
    // Transform vocabulary response format to slides
    // Input: { title, subtitle, words: [{ word, definition, translation, example? }] }
    // Output: array of slide objects for preview

    const slides = [];

    // Add title slide
    slides.push({
        type: 'Title',
        title: vocabData.title || 'Vocabulary',
        subtitle: vocabData.subtitle || '',
        content: ''
    });

    // Add a slide for each word using single-item transformer
    if (vocabData.words && Array.isArray(vocabData.words)) {
        vocabData.words.forEach(word => {
            slides.push(transformSingleVocabularySlide(word));
        });
    }

    return slides;
}

function transformGrammarToSlides(grammarData) {
    // Transform grammar response format to slides
    // Input: { title, subtitle, slides: [{ slide-title, content: { explanation, form, usage, examples } }] }
    // Output: array of slide objects for preview

    const slides = [];

    // Add title slide
    slides.push({
        type: 'Title',
        title: grammarData.title || 'Grammar',
        subtitle: grammarData.subtitle || '',
        content: ''
    });

    // Add a slide for each grammar section using single-item transformer
    if (grammarData.slides && Array.isArray(grammarData.slides)) {
        grammarData.slides.forEach(slide => {
            slides.push(transformSingleGrammarSlide(slide));
        });
    }

    return slides;
}

function transformQuizToSlides(quizData) {
    // Transform quiz response format to slides
    // Input: { title, subtitle, quiz-type, focus, questions: [{ question, options, correct-answer }] }
    // Output: array of slide objects for preview

    const slides = [];

    // Add title slide
    slides.push({
        type: 'Title',
        title: quizData.title || 'Quiz',
        subtitle: quizData.subtitle || '',
        content: `Type: ${quizData['quiz-type'] || 'Multiple Choice'} | Focus: ${quizData.focus || 'General'}`
    });

    // Add a slide for each question using single-item transformer
    if (quizData.questions && Array.isArray(quizData.questions)) {
        quizData.questions.forEach((q, index) => {
            slides.push(transformSingleQuizSlide(q, `Question ${index + 1}`));
        });
    }

    return slides;
}

function transformHomeworkToSlides(homeworkData) {
    // Transform homework response format to slides
    // Input: { title, subtitle, focus, homework-type, tasks: [{ instruction, items }] }
    // Output: array of slide objects for preview

    const slides = [];

    // Add title slide
    slides.push({
        type: 'Title',
        title: homeworkData.title || 'Homework',
        subtitle: homeworkData.subtitle || '',
        content: `Type: ${homeworkData['homework-type'] || 'Exercise'} | Focus: ${homeworkData.focus || 'General'}`
    });

    // Add a slide for each task using single-item transformer
    if (homeworkData.tasks && Array.isArray(homeworkData.tasks)) {
        homeworkData.tasks.forEach((task, index) => {
            slides.push(transformSingleHomeworkSlide(task, `Task ${index + 1}`));
        });
    }

    return slides;
}

function transformEditedSlide(slideData, originalType, existingSlide = null) {
    // Transform an edited slide response using the same single-item transformers
    // This ensures consistency between generation and edit flows
    //
    // existingSlide can be passed to preserve the original title (e.g., "Question 3")

    switch (originalType) {
        case 'vocabulary':
            return transformSingleVocabularySlide(slideData);

        case 'grammar':
            return transformSingleGrammarSlide(slideData);

        case 'quiz':
            // Preserve existing title like "Question 3" if available
            return transformSingleQuizSlide(slideData, existingSlide?.title || 'Question');

        case 'homework':
            // Preserve existing title like "Task 2" if available
            return transformSingleHomeworkSlide(slideData, existingSlide?.title || 'Task');

        default:
            // Generic fallback
            return {
                type: slideData.type || 'Content',
                title: slideData.title || '',
                subtitle: slideData.subtitle || '',
                content: slideData.content || '',
                example: slideData.example || ''
            };
    }
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
                console.log('Updating progress:', message.stage);
                updateProgressInPreviewArea(message.stage);
                break;

            case 'preview':
                showSlidePreview(message.slides, message.summary);
                break;

            case 'insertProgress':
                updateProgressInPreviewArea(
                    `Inserting slide ${message.current} of ${message.total}...`
                );
                break;

            case 'success':
                hideProgress();
                showSuccess(message.message);
                setProcessing(false);
                break;

            case 'error':
                hideProgress();
                showError(message.message);
                setProcessing(false);
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

function showProgress(status) {
    const template = document.getElementById('progressMessageTemplate');
    const clone = template.content.cloneNode(true);
    const progressEl = clone.querySelector('.message-progress');

    progressEl.querySelector('.progress-status').textContent = status;

    // Store reference for updates
    state.progressElement = progressEl;
    appendToChatBody(progressEl);
}

function updateProgress(status) {
    if (!state.progressElement) {
        showProgress(status);
        return;
    }

    state.progressElement.querySelector('.progress-status').textContent = status;
}

function hideProgress() {
    if (state.progressElement && state.progressElement.parentNode) {
        state.progressElement.remove();
    }
    state.progressElement = null;
}

function cancelGeneration() {
    console.log('Generation cancelled by user');
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

    // Attach cancel button handler
    const cancelBtn = progressEl.querySelector('.progress-cancel-btn');
    if (cancelBtn) {
        cancelBtn.addEventListener('click', cancelGeneration);
    }

    // Store reference for updates
    state.progressElement = progressEl;

    // Append to chat body (after user message)
    appendToChatBody(progressEl);
}

function updateProgressInPreviewArea(status) {
    console.log('updateProgressInPreviewArea called:', status, 'progressElement exists:', !!state.progressElement);
    if (!state.progressElement) {
        console.log('No progress element, creating new one');
        showProgressInPreviewArea(status);
        return;
    }

    state.progressElement.querySelector('.progress-status').textContent = status;
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

    // Show dismiss message (same style as AI messages)
    addAIMessage(message);

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
    setProcessing(false);  // Ready for navigation
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
    previewEl.querySelector('#previewCancelBtn').addEventListener('click', cancelPreview);
}

function cancelPreview() {
    console.log('Preview cancelled by user');
    dismissPreview('Preview cancelled.');
}

function navigateBack() {
    if (state.currentSlideIndex > 0) {
        state.currentSlideIndex--;
        // Update edit mode to new slide if still editing
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
        // Update edit mode to new slide if still editing
        if (state.isEditMode) {
            state.editingSlideIndex = state.currentSlideIndex;
            updateEditBadgeSlideNumber(state.currentSlideIndex + 1);
        }
        updateSlideDisplay();
    } else {
        // On last slide, "Done" inserts all slides
        insertAllSlides();
    }
}

function skipSlide() {
    state.skippedSlides.add(state.currentSlideIndex);
    if (state.currentSlideIndex < state.slides.length - 1) {
        state.currentSlideIndex++;
        // Update edit mode to new slide if still editing
        if (state.isEditMode) {
            state.editingSlideIndex = state.currentSlideIndex;
            updateEditBadgeSlideNumber(state.currentSlideIndex + 1);
        }
        updateSlideDisplay();
    } else {
        updateSlideDisplay();
    }
}

function editSlide() {
    // Enter edit mode for the current slide
    const { messageInput, typeSelector } = state.elements;

    // Set edit mode state
    state.isEditMode = true;
    state.editingSlideIndex = state.currentSlideIndex;

    // Change input placeholder to indicate edit mode
    messageInput.placeholder = `Describe what to change on slide ${state.currentSlideIndex + 1}...`;
    messageInput.value = '';

    // Hide type selector (not relevant during edit)
    if (typeSelector) {
        typeSelector.classList.add('hidden');
    }

    // Show edit badge
    showEditBadge(state.currentSlideIndex + 1);

    // Focus the input
    messageInput.focus();
}

function exitEditMode() {
    // Exit edit mode and restore normal state
    const { messageInput, typeSelector } = state.elements;

    // Reset edit mode state
    state.isEditMode = false;
    state.editingSlideIndex = null;

    // Restore original placeholder
    messageInput.placeholder = 'Ask Teachers Center to help you create slides...';

    // Show type selector again
    if (typeSelector) {
        typeSelector.classList.remove('hidden');
    }

    // Remove edit badge
    removeEditBadge();
}

function showEditBadge(slideNumber) {
    // Show the edit badge with the slide number
    const { editBadge, editBadgeSlideNum } = state.elements;
    if (editBadge && editBadgeSlideNum) {
        editBadgeSlideNum.textContent = slideNumber;
        editBadge.classList.remove('hidden');
    }
}

function updateEditBadgeSlideNumber(slideNumber) {
    // Update the slide number in the edit badge and placeholder (when navigating while in edit mode)
    const { editBadgeSlideNum, messageInput } = state.elements;
    if (editBadgeSlideNum) {
        editBadgeSlideNum.textContent = slideNumber;
    }
    if (messageInput) {
        messageInput.placeholder = `Describe what to change on slide ${slideNumber}...`;
    }
}

function removeEditBadge() {
    // Hide the edit badge
    const { editBadge } = state.elements;
    if (editBadge) {
        editBadge.classList.add('hidden');
    }
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

    // Exit edit mode if active
    if (state.isEditMode) {
        exitEditMode();
    }

    // Hide preview immediately
    hidePreviewArea();

    // Show progress
    setProcessing(true);
    showProgressInPreviewArea('Inserting slides...');

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

    // Escape key: exit edit mode if editing, otherwise focus input
    if (e.key === 'Escape') {
        if (state.isEditMode) {
            e.preventDefault();
            exitEditMode();
            return;
        }
    }

    // Q key: cancel/quit in both progress and preview states
    if (e.key.toLowerCase() === 'q' && document.activeElement !== state.elements.messageInput) {
        if (state.isProcessing) {
            e.preventDefault();
            cancelGeneration();
            return;
        }
        if (state.isInPreviewMode) {
            e.preventDefault();
            cancelPreview();
            return;
        }
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

function setProcessing(isProcessing) {
    state.isProcessing = isProcessing;
    const { messageInput, typeOptions } = state.elements;

    // Disable/enable input and type buttons during processing
    messageInput.disabled = isProcessing;
    typeOptions.querySelectorAll('.type-option').forEach(btn => {
        btn.disabled = isProcessing;
    });

    if (isProcessing) {
        messageInput.placeholder = 'Generating content...';
    } else {
        // Restore placeholder based on selected type
        const selectedBtn = typeOptions.querySelector('.type-option.selected');
        messageInput.placeholder = selectedBtn?.dataset.placeholder || 'Type your request...';
    }
}

function appendToChatBody(element) {
    const { chatBody, welcomeState } = state.elements;

    // Make sure welcome state is hidden when adding messages
    welcomeState.classList.add('hidden');

    chatBody.appendChild(element);

    // Scroll to bottom
    chatBody.scrollTop = chatBody.scrollHeight;
}

function insertBeforePreview(element) {
    // Insert an element before the preview container (used for edit messages)
    // This keeps the preview visible after edits
    const { chatBody, welcomeState } = state.elements;

    welcomeState.classList.add('hidden');

    if (state.previewElement && state.previewElement.parentNode === chatBody) {
        chatBody.insertBefore(element, state.previewElement);
    } else {
        // Fallback to append if no preview
        chatBody.appendChild(element);
    }

    // Scroll to show preview
    if (state.previewElement) {
        state.previewElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
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
