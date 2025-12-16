/* global document, Office, PowerPoint */

// ============================================================================
// WEBSOCKET CONNECTION CONFIGURATION
// ============================================================================

const WS_BACKEND_URL = 'ws://localhost:2000/ws';
const USER_ID = 123; // HARDCODED for Phase 1
const CHANNEL_NAME = "powerpoint-session-001"; // HARDCODED for Phase 1
const MAX_RECONNECT_ATTEMPTS = 5;

let wsConnection = null;
let wsConnectionState = 'disconnected'; // 'disconnected' | 'connecting' | 'connected'
let reconnectAttempts = 0;

// ============================================================================
// APPLICATION STATE
// ============================================================================

// Current selected category
let currentCategory = 'vocabulary';

// Class context settings
let classContext = {
  language: 'English',
  level: 'B1',
  className: 'Class 7A',
  nativeLanguage: '',
  ageGroup: ''
};

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    loadSettings();
    initializeUI();
    updateContextDisplay();
    connectWebSocket();
  }
});

// ============================================================================
// WEBSOCKET CONNECTION MANAGEMENT
// ============================================================================

function connectWebSocket() {
  if (wsConnectionState === 'connecting' || wsConnectionState === 'connected') {
    console.log('WebSocket already connecting or connected');
    return;
  }

  console.log(`Connecting to WebSocket at ${WS_BACKEND_URL}...`);
  wsConnectionState = 'connecting';
  updateConnectionStatus();

  try {
    wsConnection = new WebSocket(WS_BACKEND_URL);

    wsConnection.onopen = function(event) {
      console.log('WebSocket connected successfully');
      wsConnectionState = 'connected';
      reconnectAttempts = 0;
      updateConnectionStatus();
      addMessage('Connected to server', 'ai');
    };

    wsConnection.onmessage = handleWebSocketMessage;
    wsConnection.onerror = handleWebSocketError;
    wsConnection.onclose = handleWebSocketClose;

  } catch (error) {
    console.error('Failed to create WebSocket connection:', error);
    wsConnectionState = 'disconnected';
    updateConnectionStatus();
    addMessage('Failed to connect to server', 'ai');
  }
}

function disconnectWebSocket() {
  if (wsConnection) {
    console.log('Disconnecting WebSocket...');
    wsConnection.close();
    wsConnection = null;
    wsConnectionState = 'disconnected';
    updateConnectionStatus();
  }
}

function handleWebSocketMessage(event) {
  console.log('Received WebSocket message:', event.data);

  try {
    const response = JSON.parse(event.data);
    console.log('Parsed response:', response);

    // Remove loading indicator
    removeLoadingMessage();

    // Display in chat
    displayResponseInChat(response);

    // Insert into slide (will implement later)
    // insertResponseIntoSlide(response);

  } catch (error) {
    console.error('Failed to parse WebSocket message:', error);
    removeLoadingMessage();
    addMessage('Error: Failed to process response', 'ai');
  }
}

function handleWebSocketError(error) {
  console.error('WebSocket error:', error);
  addMessage('Connection error occurred', 'ai');
}

function handleWebSocketClose(event) {
  console.log('WebSocket closed:', event.code, event.reason);
  wsConnectionState = 'disconnected';
  updateConnectionStatus();

  // Attempt to reconnect if not a clean close
  if (event.code !== 1000 && reconnectAttempts < MAX_RECONNECT_ATTEMPTS) {
    reconnectAttempts++;
    const delay = Math.min(1000 * Math.pow(2, reconnectAttempts), 30000);
    console.log(`Attempting to reconnect in ${delay}ms (attempt ${reconnectAttempts}/${MAX_RECONNECT_ATTEMPTS})`);

    setTimeout(() => {
      addMessage(`Reconnecting... (attempt ${reconnectAttempts}/${MAX_RECONNECT_ATTEMPTS})`, 'ai');
      connectWebSocket();
    }, delay);
  } else if (reconnectAttempts >= MAX_RECONNECT_ATTEMPTS) {
    addMessage('Connection lost. Please refresh the page to reconnect.', 'ai');
  }
}

function sendWebSocketMessage(payload) {
  if (wsConnectionState !== 'connected') {
    console.error('Cannot send message: WebSocket not connected');
    addMessage('Error: Not connected to server. Please wait...', 'ai');
    return false;
  }

  try {
    const jsonString = JSON.stringify(payload);
    console.log('Sending WebSocket message:', jsonString);
    wsConnection.send(jsonString);
    return true;
  } catch (error) {
    console.error('Failed to send WebSocket message:', error);
    addMessage('Error: Failed to send message', 'ai');
    return false;
  }
}

function updateConnectionStatus() {
  // This will be implemented when we add the UI status indicator
  console.log(`Connection status: ${wsConnectionState}`);

  // For now, just update a class on the body element
  const body = document.body;
  if (body) {
    body.classList.remove('ws-disconnected', 'ws-connecting', 'ws-connected');
    body.classList.add(`ws-${wsConnectionState}`);
  }
}

function displayResponseInChat(response) {
  // Check if this is an error response
  if (response.error) {
    addMessage(`Error: ${response.error}`, 'ai');
    return;
  }

  // Check if this is a vocabulary response
  if (response.title && response.words && Array.isArray(response.words)) {
    displayVocabularyResponse(response);
    return;
  }

  // Generic response display
  addMessage(JSON.stringify(response, null, 2), 'ai');
}

function displayVocabularyResponse(response) {
  let chatMessage = `${response.title}\n${response.subtitle}\n\n`;

  response.words.forEach((word, index) => {
    chatMessage += `${index + 1}. ${word.word} - ${word.translation}\n`;
    chatMessage += `   ${word.definition}\n\n`;
  });

  addMessage(chatMessage, 'ai');
}

function addLoadingMessage() {
  const chatBody = document.getElementById('body');

  const messageDiv = document.createElement('div');
  messageDiv.className = 'message ai-message loading-message';

  const avatarDiv = document.createElement('div');
  avatarDiv.className = 'message-avatar';
  avatarDiv.innerHTML = '<span class="material-icons">psychology</span>';

  const contentDiv = document.createElement('div');
  contentDiv.className = 'message-content';

  const textDiv = document.createElement('div');
  textDiv.className = 'message-text';
  textDiv.textContent = 'Generating content...';

  contentDiv.appendChild(textDiv);
  messageDiv.appendChild(avatarDiv);
  messageDiv.appendChild(contentDiv);

  chatBody.appendChild(messageDiv);
  chatBody.scrollTop = chatBody.scrollHeight;
}

function removeLoadingMessage() {
  const loadingMsg = document.querySelector('.loading-message');
  if (loadingMsg) {
    loadingMsg.remove();
  }
}

// ============================================================================
// UI INITIALIZATION
// ============================================================================

function initializeUI() {
  // Initialize category chips
  const chips = document.querySelectorAll('.chip');
  chips.forEach(chip => {
    chip.addEventListener('click', handleCategoryClick);
  });

  // Initialize send button
  const sendButton = document.querySelector('.send-button');
  if (sendButton) {
    sendButton.addEventListener('click', handleSendMessage);
  }

  // Initialize input field (Enter key to send)
  const messageInput = document.querySelector('.message-input');
  if (messageInput) {
    messageInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        handleSendMessage();
      }
    });
  }

  // Initialize resource button (placeholder for now)
  const resourceButton = document.querySelector('.resource-button');
  if (resourceButton) {
    resourceButton.addEventListener('click', handleResourceClick);
  }

  // Initialize settings button
  const settingsBtn = document.getElementById('settingsBtn');
  if (settingsBtn) {
    settingsBtn.addEventListener('click', openSettingsModal);
  }

  // Initialize modal close buttons
  const closeModalBtn = document.getElementById('closeModalBtn');
  const cancelSettingsBtn = document.getElementById('cancelSettingsBtn');
  if (closeModalBtn) {
    closeModalBtn.addEventListener('click', closeSettingsModal);
  }
  if (cancelSettingsBtn) {
    cancelSettingsBtn.addEventListener('click', closeSettingsModal);
  }

  // Initialize save settings button
  const saveSettingsBtn = document.getElementById('saveSettingsBtn');
  if (saveSettingsBtn) {
    saveSettingsBtn.addEventListener('click', saveSettings);
  }

  // Close modal when clicking outside
  const modal = document.getElementById('settingsModal');
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        closeSettingsModal();
      }
    });
  }

  console.log('Chat UI initialized successfully');
}

function handleCategoryClick(event) {
  // Get the chip that was clicked
  const clickedChip = event.currentTarget;

  // Remove active class from all chips
  document.querySelectorAll('.chip').forEach(chip => {
    chip.classList.remove('active');
  });

  // Add active class to clicked chip
  clickedChip.classList.add('active');

  // Update current category
  currentCategory = clickedChip.dataset.category;

  console.log(`Switched to category: ${currentCategory}`);

  // Add system message to chat
  addMessage(`Switched to ${currentCategory} mode. How can I help you?`, 'ai');
}

function handleSendMessage() {
  const messageInput = document.querySelector('.message-input');
  const messageText = messageInput.value.trim();

  if (messageText === '') {
    return;
  }

  // Add user message to chat
  addMessage(messageText, 'user');

  // Clear input
  messageInput.value = '';

  // Simulate AI response (placeholder)
  setTimeout(() => {
    const dummyResponse = getDummyResponse(currentCategory, messageText);
    addMessage(dummyResponse, 'ai');
  }, 800);
}

function handleResourceClick() {
  // Placeholder for resource attachment
  console.log('Resource button clicked - functionality to be implemented');
  showStatus('Resource attachment feature coming soon!', true);
}

function addMessage(text, type) {
  const chatBody = document.getElementById('body');

  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${type}-message`;

  const avatarDiv = document.createElement('div');
  avatarDiv.className = 'message-avatar';
  avatarDiv.innerHTML = `<span class="material-icons">${type === 'ai' ? 'smart_toy' : 'person'}</span>`;

  const contentDiv = document.createElement('div');
  contentDiv.className = 'message-content';

  const textDiv = document.createElement('div');
  textDiv.className = 'message-text';
  textDiv.textContent = text;

  const timeDiv = document.createElement('div');
  timeDiv.className = 'message-time';
  timeDiv.textContent = 'Just now';

  contentDiv.appendChild(textDiv);
  contentDiv.appendChild(timeDiv);

  if (type === 'ai') {
    messageDiv.appendChild(avatarDiv);
    messageDiv.appendChild(contentDiv);
  } else {
    messageDiv.appendChild(contentDiv);
    messageDiv.appendChild(avatarDiv);
  }

  chatBody.appendChild(messageDiv);

  // Scroll to bottom
  chatBody.scrollTop = chatBody.scrollHeight;
}

function getDummyResponse(category, userMessage) {
  const responses = {
    vocabulary: "I can help you create vocabulary lists! Just tell me the topic and student level, and I'll generate engaging vocabulary content for your slides.",
    quizzes: "Let's create a quiz! What topic would you like to focus on? I can generate multiple choice, true/false, or fill-in-the-blank questions.",
    homework: "I'll help you create homework assignments. What subject and level are you teaching?",
    grammar: "Grammar is my specialty! Which grammar concept would you like to cover? Tenses, articles, conditionals?",
    reading: "I can create reading comprehension materials. What reading level and topic are you interested in?",
    listening: "For listening activities, I can help you design comprehension exercises. What's your focus area?"
  };

  return responses[category] || "How can I help you with your teaching materials?";
}

function showStatus(message, isSuccess = true) {
  console.log(`Status: ${message} (${isSuccess ? 'success' : 'error'})`);
  // Add a temporary toast notification
  addMessage(message, 'ai');
}

// ============================================================================
// SETTINGS MANAGEMENT
// ============================================================================

function loadSettings() {
  try {
    const savedSettings = localStorage.getItem('classContext');
    if (savedSettings) {
      const parsed = JSON.parse(savedSettings);
      classContext = { ...classContext, ...parsed };
      console.log('Settings loaded from localStorage:', classContext);
    } else {
      console.log('No saved settings found, using defaults');
    }
  } catch (error) {
    console.error('Error loading settings:', error);
  }
}

function saveSettings() {
  try {
    // Get values from form
    classContext.language = document.getElementById('settingsLanguage').value;
    classContext.level = document.getElementById('settingsLevel').value;
    classContext.className = document.getElementById('settingsClassName').value.trim();
    classContext.nativeLanguage = document.getElementById('settingsNativeLanguage').value;
    classContext.ageGroup = document.getElementById('settingsAgeGroup').value.trim();

    // Validate
    if (!classContext.className) {
      addMessage('Please enter a class name.', 'ai');
      return;
    }

    // Save to localStorage
    localStorage.setItem('classContext', JSON.stringify(classContext));
    console.log('Settings saved:', classContext);

    // Update display
    updateContextDisplay();

    // Close modal
    closeSettingsModal();

    // Notify user
    addMessage(`Settings updated: ${classContext.level} ${classContext.language} • ${classContext.className}`, 'ai');
  } catch (error) {
    console.error('Error saving settings:', error);
    addMessage('Error saving settings. Please try again.', 'ai');
  }
}

function openSettingsModal() {
  // Populate form with current values
  document.getElementById('settingsLanguage').value = classContext.language;
  document.getElementById('settingsLevel').value = classContext.level;
  document.getElementById('settingsClassName').value = classContext.className;
  document.getElementById('settingsNativeLanguage').value = classContext.nativeLanguage || '';
  document.getElementById('settingsAgeGroup').value = classContext.ageGroup || '';

  // Show modal
  const modal = document.getElementById('settingsModal');
  if (modal) {
    modal.style.display = 'flex';
  }
}

function closeSettingsModal() {
  const modal = document.getElementById('settingsModal');
  if (modal) {
    modal.style.display = 'none';
  }
}

function updateContextDisplay() {
  const display = document.getElementById('contextDisplay');
  if (display) {
    display.textContent = `${classContext.level} ${classContext.language} • ${classContext.className}`;
  }
}

function getClassContext() {
  return classContext;
}

// ============================================================================
// OLD FUNCTIONS (Kept for future reference - can be removed if not needed)
// ============================================================================

/*
// These functions were part of the old vocabulary builder form interface
// They are kept here for reference but are not currently used

function updateSmartDefaultsDisplay() { ... }
function toggleAdvancedOptions() { ... }
function validateForm() { ... }
function getFormData() { ... }
async function generateVocabulary() { ... }
function insertTextFallback(result) { ... }
async function insertSlidesFromResult(result) { ... }
*/
