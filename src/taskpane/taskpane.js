/* global Office, PowerPoint */

/**
 * Teacher Assistant Taskpane - Full chat/preview UI with direct PowerPoint API access.
 * Ported from dialog.js to run in the taskpane sidebar (~350px).
 * No parent-child messaging needed — calls PowerPoint.run() directly.
 */

// Import CSS for webpack bundling
import './taskpane.css';

// ============================================
// WEBSOCKET CONFIGURATION
// ============================================

const WS_URL = process.env.WS_URL;
const USER_ID = 'user-123';  // TODO: implement proper user management
const CHANNEL_NAME = 'powerpoint-taskpane';
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
    isInPreviewMode: false,

    // Edit mode state
    originalRequest: null,
    editingSlideIndex: null,
    isEditMode: false,

    // Progress state
    progressElement: null,
    cancelled: false,
    staleResponses: 0,

    // Deduplication
    lastMessageId: null,

    // Type selection
    selectedType: 'vocabulary',

    // Settings
    settingsConfirmed: false,
    settings: {
        language: 'English',
        level: 'B1',
        nativeLanguage: 'No',
        ageGroup: ''
    },

    // Current file path (for per-file settings)
    currentFilePath: null,

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

// ============================================
// INITIALIZATION
// ============================================

// ============================================
// OFFICE INITIALIZATION
// ============================================

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById('sideload-msg').style.display = 'none';
        document.getElementById('app-body').style.display = 'flex';
        initializeTaskpane();
    }
});

function initializeTaskpane() {
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
        settingsModal: document.getElementById('settingsModal'),
        closeModalBtn: document.getElementById('closeModalBtn'),
        cancelSettingsBtn: document.getElementById('cancelSettingsBtn'),
        saveSettingsBtn: document.getElementById('saveSettingsBtn'),
        settingsLanguage: document.getElementById('settingsLanguage'),
        settingsLevel: document.getElementById('settingsLevel'),
        settingsNativeLanguage: document.getElementById('settingsNativeLanguage'),
        settingsAgeGroup: document.getElementById('settingsAgeGroup')
    };

    loadSettings();
    setupEventListeners();

    // Connect to WebSocket backend
    setTimeout(() => {
        connectWebSocket();
    }, 500);

    console.log('Taskpane initialized');
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

    // Auto-resize textarea
    messageInput.addEventListener('input', () => {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.min(messageInput.scrollHeight, 100) + 'px';
    });

    // Header buttons
    newChatBtn.addEventListener('click', handleNewChat);

    // Edit badge cancel
    const { editBadgeCancelBtn } = state.elements;
    if (editBadgeCancelBtn) {
        editBadgeCancelBtn.addEventListener('click', exitEditMode);
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

    typeOptions.querySelectorAll('.type-option').forEach(opt => opt.classList.remove('selected'));
    btn.classList.add('selected');
    state.selectedType = type;
    messageInput.placeholder = placeholder;
    clearTypeValidationError();
    messageInput.focus();
}

function resetTypeSelection() {
    const { typeOptions, messageInput } = state.elements;
    typeOptions.querySelectorAll('.type-option').forEach(opt => opt.classList.remove('selected'));
    const defaultBtn = typeOptions.querySelector('[data-type="vocabulary"]');
    if (defaultBtn) {
        defaultBtn.classList.add('selected');
        state.selectedType = 'vocabulary';
        messageInput.placeholder = defaultBtn.dataset.placeholder || 'Type your request...';
    }
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
    if (!state.settingsConfirmed) {
        openSettingsModal();
        return;
    }

    const { messageInput } = state.elements;
    const content = messageInput.value.trim();
    if (!content || state.isProcessing) return;

    // Edit mode
    if (state.isEditMode && state.editingSlideIndex !== null) {
        handleEditSend(content);
        return;
    }

    // if (!validateTypeSelection()) return;

    // Dismiss existing preview
    if (state.isInPreviewMode && state.slides.length > 0) {
        const count = state.slides.length;
        dismissPreview(`${count} slide${count !== 1 ? 's' : ''} not inserted`);
    }

    state.lastMessageId = null;
    state.cancelled = false;
    state.elements.welcomeState.classList.add('hidden');

    addUserMessage(content);
    messageInput.value = '';
    messageInput.style.height = 'auto';

    setProcessing(true);
    showProgressInPreviewArea('Generating content...');
    state.originalRequest = content;

    const sent = sendWebSocketMessage({
        type: state.selectedType,
        content: content
    });

    if (!sent) connectWebSocket();
}

function handleEditSend(editInstruction) {
    const { messageInput } = state.elements;
    const slideIndex = state.editingSlideIndex;
    const currentSlide = state.slides[slideIndex];

    // Show user message before preview
    const messageContent = `Edit slide ${slideIndex + 1}: ${editInstruction}`;
    const template = document.getElementById('userMessageTemplate');
    const clone = template.content.cloneNode(true);
    const messageEl = clone.querySelector('.message-user');
    messageEl.textContent = messageContent;
    insertBeforePreview(messageEl);
    state.messages.push({ type: 'user', content: messageContent });

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
            originalRequest: state.originalRequest,
            originalType: state.selectedType
        }
    });

    if (!sent) connectWebSocket();
}

function handleNewChat() {
    state.messages = [];
    state.pendingRequest = null;
    state.currentSlideIndex = 0;
    setProcessing(false);
    state.conversationId = null;
    state.originalRequest = null;
    state.editingSlideIndex = null;
    state.isEditMode = false;

    hidePreviewArea();

    const { chatBody, welcomeState } = state.elements;
    chatBody.innerHTML = '';
    chatBody.appendChild(welcomeState);
    welcomeState.classList.remove('hidden');
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
        if (!state.conversationId) {
            state.conversationId = `conv-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
        }

        const wsMessage = {
            'user-id': USER_ID,
            'channel-name': CHANNEL_NAME,
            'conversation-id': state.conversationId,
            type: message.type,
            content: message.content,
            requirements: {
                language: state.settings.language,
                level: state.settings.level,
                'native-language': state.settings.nativeLanguage || 'No',
                'age-group': state.settings.ageGroup || null
            },
            ...(message.edit && { edit: message.edit })
        };

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

        // Ignore stale responses
        if (state.staleResponses > 0) {
            if (message.type !== 'progress') state.staleResponses--;
            return;
        }

        if (message['requirements-not-met']) {
            hideProgress();
            addAIMessage(message['requirements-not-met']);
            setProcessing(false);
            return;
        }

        if (message.type === 'progress') {
            updateProgressInPreviewArea(message.stage || message.message);
            return;
        }

        // Edit response
        if (message.type === 'edit' && message.edit) {
            const slideIndex = message.edit.slideIndex;
            const existingSlide = state.slides[slideIndex];
            const transformedSlide = transformEditedSlide(message.edit.slide, state.selectedType, existingSlide);

            if (transformedSlide && slideIndex >= 0 && slideIndex < state.slides.length) {
                state.slides[slideIndex] = transformedSlide;
                state.currentSlideIndex = slideIndex;
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

        // NEW — Handle unified conversation response (slide-title + content schema)
        if (message.slides) {
            const slides = transformConversationResponse(message);
            if (slides && slides.length > 0) {
                showSlidePreview(slides, message.title || 'Generated Content');
            } else {
                showError('No slides generated. Please try a different request.');
                setProcessing(false);
            }
            return;
        }

        // OLD — Route by type (vocabulary, grammar, quiz, homework) — commented out
        // if (message.type) {
        //     const slides = transformResponseByType(message);
        //     if (slides && slides.length > 0) {
        //         const titles = { vocabulary: 'Vocabulary', grammar: 'Grammar', quiz: 'Quiz', homework: 'Homework' };
        //         showSlidePreview(slides, message.title || titles[message.type] || 'Generated Content');
        //     } else {
        //         showError(`No ${message.type} content generated. Please try a different request.`);
        //         setProcessing(false);
        //     }
        //     return;
        // }

        // OLD — Legacy fallback — commented out
        // if (message.slides || message.data) {
        //     const slides = transformBackendSlides(message.slides || message.data);
        //     if (slides && slides.length > 0) {
        //         showSlidePreview(slides, message.summary || 'Generated Slides');
        //     } else {
        //         showError('No slides generated. Please try a different request.');
        //         setProcessing(false);
        //     }
        //     return;
        // }

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
// TRANSFORM FUNCTIONS
// ============================================

function transformResponseByType(message) {
    switch (message.type) {
        case 'vocabulary': return transformVocabularyToSlides(message);
        case 'grammar': return transformGrammarToSlides(message);
        case 'quiz': return transformQuizToSlides(message);
        case 'homework': return transformHomeworkToSlides(message);
        default: return transformBackendSlides(message.slides || message.data || []);
    }
}

function transformBackendSlides(backendSlides) {
    if (!backendSlides || !Array.isArray(backendSlides)) return [];
    return backendSlides.map(slide => ({
        type: slide.type || 'Content',
        title: slide.title || '',
        subtitle: slide.subtitle || '',
        content: slide.content || slide.body || '',
        example: slide.example || slide['example-sentence'] || ''
    }));
}

function transformConversationResponse(message) {
    // Transforms the unified conversation-content.edn response schema:
    // { title, subtitle, slides: [{ "slide-title", content }] }
    const slides = [];

    if (message.title) {
        slides.push({
            type: 'Title',
            title: message.title,
            subtitle: message.subtitle || '',
            content: '',
            example: ''
        });
    }

    if (message.slides && Array.isArray(message.slides)) {
        message.slides.forEach(slide => {
            slides.push({
                type: 'Content',
                title: slide['slide-title'] || slide.title || '',
                subtitle: slide.subtitle || '',
                content: slide.content || '',
                example: slide.example || ''
            });
        });
    }

    return slides;
}

// Single-item transforms
function transformSingleVocabularySlide(wordData) {
    return {
        type: 'Vocabulary',
        title: wordData.word || wordData.title || '',
        subtitle: wordData.translation || wordData.subtitle || '',
        content: wordData.definition || wordData.content || '',
        example: wordData.example || ''
    };
}

function transformSingleGrammarSlide(slideData) {
    let contentText = '';
    if (typeof slideData.content === 'string') {
        contentText = slideData.content;
    } else if (typeof slideData.content === 'object' && slideData.content !== null) {
        contentText = slideData.content.explanation || '';
    }

    let exampleText = '';
    const examples = slideData.examples || (slideData.content && slideData.content.examples) || [];
    if (Array.isArray(examples)) {
        examples.forEach(ex => {
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
    const groupedQuestions = questionData['slide-questions'] || questionData.slideQuestions;
    if (groupedQuestions && Array.isArray(groupedQuestions)) {
        let contentText = '';
        groupedQuestions.forEach((q, i) => {
            contentText += `${i + 1}. ${q.question || ''}\n`;
            if (q.options && Array.isArray(q.options)) {
                q.options.forEach((opt, j) => {
                    contentText += `   ${String.fromCharCode(65 + j)}. ${opt}\n`;
                });
            }
            contentText += '\n';
        });
        return { type: 'Quiz', title: questionData.title || title, subtitle: '', content: contentText.trim(), example: '' };
    }

    let contentText = (questionData.question || '') + '\n\n';
    if (questionData.options && Array.isArray(questionData.options)) {
        questionData.options.forEach((opt, i) => {
            contentText += `${String.fromCharCode(65 + i)}. ${opt}\n`;
        });
    }
    return { type: 'Quiz', title: questionData.title || title, subtitle: '', content: contentText.trim(), example: '' };
}

function transformSingleHomeworkSlide(taskData, title = 'Task') {
    let contentText = '';
    if (taskData.instruction) contentText += `${taskData.instruction}\n\n`;
    if (taskData.items && Array.isArray(taskData.items)) {
        taskData.items.forEach((item, i) => { contentText += `${i + 1}. ${item}\n`; });
    }
    return { type: 'Homework', title: taskData.title || title, subtitle: taskData.instruction || '', content: contentText.trim(), example: '' };
}

// Multi-slide transforms
function transformVocabularyToSlides(vocabData) {
    const slides = [{ type: 'Title', title: vocabData.title || 'Vocabulary', subtitle: vocabData.subtitle || '', content: '' }];
    if (vocabData.words && Array.isArray(vocabData.words)) {
        vocabData.words.forEach(word => slides.push(transformSingleVocabularySlide(word)));
    }
    return slides;
}

function transformGrammarToSlides(grammarData) {
    const slides = [{ type: 'Title', title: grammarData.title || 'Grammar', subtitle: grammarData.subtitle || '', content: '' }];
    if (grammarData.slides && Array.isArray(grammarData.slides)) {
        grammarData.slides.forEach(slide => slides.push(transformSingleGrammarSlide(slide)));
    }
    return slides;
}

function transformQuizToSlides(quizData) {
    const slides = [{
        type: 'Title',
        title: quizData.title || 'Quiz',
        subtitle: quizData.subtitle || '',
        content: `Type: ${quizData['quiz-type'] || 'Multiple Choice'} | Focus: ${quizData.focus || 'General'}`
    }];
    if (quizData.questions && Array.isArray(quizData.questions)) {
        let questionNum = 0;
        quizData.questions.forEach(q => {
            const grouped = q['slide-questions'] || q.slideQuestions;
            if (grouped && Array.isArray(grouped)) {
                const from = questionNum + 1;
                questionNum += grouped.length;
                slides.push(transformSingleQuizSlide(q, `Questions ${from}–${questionNum}`));
            } else if (q.question) {
                questionNum++;
                slides.push(transformSingleQuizSlide(q, `Question ${questionNum}`));
            }
        });
    }
    return slides;
}

function transformHomeworkToSlides(homeworkData) {
    const slides = [{
        type: 'Title',
        title: homeworkData.title || 'Homework',
        subtitle: homeworkData.subtitle || '',
        content: `Type: ${homeworkData['homework-type'] || 'Exercise'} | Focus: ${homeworkData.focus || 'General'}`
    }];
    if (homeworkData.tasks && Array.isArray(homeworkData.tasks)) {
        homeworkData.tasks.forEach((task, index) => slides.push(transformSingleHomeworkSlide(task, `Task ${index + 1}`)));
    }
    return slides;
}

function transformEditedSlide(slideData, originalType, existingSlide = null) {
    switch (originalType) {
        case 'vocabulary': return transformSingleVocabularySlide(slideData);
        case 'grammar': return transformSingleGrammarSlide(slideData);
        case 'quiz': return transformSingleQuizSlide(slideData, existingSlide?.title || 'Question');
        case 'homework': return transformSingleHomeworkSlide(slideData, existingSlide?.title || 'Task');
        default:
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
// UI - Messages
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
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
}

function dismissPreview(message) {
    if (state.previewElement && state.previewElement.parentNode) {
        state.previewElement.remove();
    }
    addAIMessage(message);
    state.slides = [];
    state.previewElement = null;
    state.isInPreviewMode = false;
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
    previewEl.querySelector('.slide-type-badge').textContent = slide.type || 'Content';
    previewEl.querySelector('.slide-card-title').textContent = slide.title || '';
    previewEl.querySelector('.slide-card-subtitle').textContent = slide.subtitle || '';
    previewEl.querySelector('.slide-card-content').textContent = slide.content || '';

    const exampleEl = previewEl.querySelector('.slide-card-example');
    if (slide.example) {
        exampleEl.textContent = slide.example;
        exampleEl.classList.remove('hidden');
    } else {
        exampleEl.classList.add('hidden');
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
    if (state.slides.length <= 1) {
        state.slides.splice(0, 1);
        dismissPreview('All slides removed.');
        return;
    }

    state.slides.splice(state.currentSlideIndex, 1);
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
    const { messageInput, typeSelector } = state.elements;
    state.isEditMode = true;
    state.editingSlideIndex = state.currentSlideIndex;
    messageInput.placeholder = `Describe what to change on slide ${state.currentSlideIndex + 1}...`;
    messageInput.value = '';
    if (typeSelector) typeSelector.classList.add('hidden');
    showEditBadge(state.currentSlideIndex + 1);
    messageInput.focus();
}

function exitEditMode() {
    const { messageInput, typeSelector } = state.elements;
    state.isEditMode = false;
    state.editingSlideIndex = null;
    messageInput.placeholder = 'Type your request...';
    if (typeSelector) typeSelector.classList.remove('hidden');
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

// ============================================
// POWERPOINT SLIDE INSERTION (Direct API)
// ============================================

async function insertAllSlides() {
    if (state.slides.length === 0) {
        showError('No slides to insert.');
        return;
    }

    const slidesToInsert = [...state.slides];

    if (state.isEditMode) exitEditMode();
    hidePreviewArea();

    setProcessing(true);
    showProgressInPreviewArea('Inserting slides...');

    try {
        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;

            for (let i = 0; i < slidesToInsert.length; i++) {
                const slideData = slidesToInsert[i];

                updateProgressInPreviewArea(`Inserting slide ${i + 1} of ${slidesToInsert.length}...`);

                presentation.slides.add();
                await context.sync();

                presentation.slides.load('items');
                await context.sync();

                const slide = presentation.slides.items[presentation.slides.items.length - 1];

                slide.shapes.load('items');
                await context.sync();

                const shapesToDelete = slide.shapes.items.slice();
                for (const shape of shapesToDelete) {
                    shape.delete();
                }
                await context.sync();

                await createSlideContent(slide, slideData, context);
            }

            await context.sync();
        });

        hideProgress();
        showSuccess(`${slidesToInsert.length} slide${slidesToInsert.length !== 1 ? 's' : ''} inserted successfully`);
        setProcessing(false);
    } catch (error) {
        console.error('Error inserting slides:', error);
        hideProgress();
        showError(`Failed to insert slides: ${error.message}`);
        setProcessing(false);
    }
}

async function createSlideContent(slide, slideData, context) {
    const isTitle = slideData.type === 'Title';

    // Title
    const titleShape = slide.shapes.addTextBox(slideData.title || '');
    titleShape.left = 50;
    titleShape.top = isTitle ? 180 : 40;
    titleShape.width = 620;
    titleShape.height = isTitle ? 80 : 60;
    await context.sync();

    titleShape.textFrame.textRange.font.bold = true;
    titleShape.textFrame.textRange.font.size = isTitle ? 44 : 32;
    titleShape.textFrame.textRange.font.color = '#d13438';
    if (isTitle) titleShape.textFrame.horizontalAlignment = 'Center';
    await context.sync();

    // Subtitle
    if (slideData.subtitle) {
        const subtitleShape = slide.shapes.addTextBox(slideData.subtitle);
        subtitleShape.left = 50;
        subtitleShape.top = isTitle ? 270 : 100;
        subtitleShape.width = 620;
        subtitleShape.height = 40;
        await context.sync();

        subtitleShape.textFrame.textRange.font.size = isTitle ? 24 : 20;
        subtitleShape.textFrame.textRange.font.color = '#605e5c';
        if (isTitle) subtitleShape.textFrame.horizontalAlignment = 'Center';
        await context.sync();
    }

    // Content
    if (slideData.content && !isTitle) {
        const contentShape = slide.shapes.addTextBox(slideData.content);
        contentShape.left = 50;
        contentShape.top = 160;
        contentShape.width = 620;
        contentShape.height = 100;
        await context.sync();

        contentShape.textFrame.textRange.font.size = 18;
        contentShape.textFrame.textRange.font.color = '#323130';
        await context.sync();
    }

    // Example
    if (slideData.example && !isTitle) {
        const exampleShape = slide.shapes.addTextBox(slideData.example);
        exampleShape.left = 50;
        exampleShape.top = 280;
        exampleShape.width = 620;
        exampleShape.height = 60;
        await context.sync();

        exampleShape.textFrame.textRange.font.size = 16;
        exampleShape.textFrame.textRange.font.italic = true;
        exampleShape.textFrame.textRange.font.color = '#605e5c';
        await context.sync();
    }
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
    appendToChatBody(successEl);
}

function showError(message) {
    hidePreviewArea();
    const template = document.getElementById('errorMessageTemplate');
    const clone = template.content.cloneNode(true);
    const errorEl = clone.querySelector('.message-error');
    errorEl.querySelector('.error-text').textContent = message;
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
        case 'r': e.preventDefault(); flashButton('navSkipBtn'); removeSlide(); break;
        case 'e': e.preventDefault(); flashButton('navEditBtn'); editSlide(); break;
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
        state.settings = { language: 'English', level: 'B1', nativeLanguage: 'No', ageGroup: '' };
        state.settingsConfirmed = false;
        setTimeout(() => openSettingsModal(), 100);
    }

    loadSettingsToForm();
    updateContextBadge();
}

function loadSettingsToForm() {
    const { settingsLanguage, settingsLevel, settingsNativeLanguage, settingsAgeGroup } = state.elements;
    settingsLanguage.value = state.settings.language;
    settingsLevel.value = state.settings.level;
    settingsNativeLanguage.value = state.settings.nativeLanguage || 'No';
    settingsAgeGroup.value = state.settings.ageGroup;
}

function saveSettings() {
    const { settingsLanguage, settingsLevel, settingsNativeLanguage, settingsAgeGroup } = state.elements;

    state.settings = {
        language: settingsLanguage.value,
        level: settingsLevel.value,
        nativeLanguage: settingsNativeLanguage.value,
        ageGroup: settingsAgeGroup.value
    };

    const storageKey = getSettingsStorageKey();
    localStorage.setItem(storageKey, JSON.stringify(state.settings));
    state.settingsConfirmed = true;
    updateContextBadge();
    closeSettingsModal();
}

function updateContextBadge() {
    const { contextBadge } = state.elements;
    if (contextBadge) {
        contextBadge.textContent = `${state.settings.level} ${state.settings.language}`;
    }
}

// ============================================
// UTILITY
// ============================================

function setProcessing(isProcessing) {
    state.isProcessing = isProcessing;
    const { messageInput } = state.elements;

    messageInput.disabled = isProcessing;
    messageInput.placeholder = isProcessing ? 'Generating content...' : 'Type your request...';
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
