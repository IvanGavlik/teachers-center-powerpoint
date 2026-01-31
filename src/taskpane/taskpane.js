/* global document, Office, PowerPoint */

// ============================================================================
// WEBSOCKET CONNECTION CONFIGURATION
// ============================================================================

// URL injected by webpack based on environment (dev/prod)
const WS_BACKEND_URL = process.env.WS_URL;
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

// Dialog reference
let dialog = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Initialize launcher view
    initializeLauncher();

    // Keep original initialization for chat view (if we switch to it)
    loadSettings();
    initializeUI();
    updateContextDisplay();
    // Don't auto-connect WebSocket - will connect when dialog opens
    // connectWebSocket();
  }
});

// ============================================================================
// DIALOG MANAGEMENT
// ============================================================================

function initializeLauncher() {
  const openDialogBtn = document.getElementById('openDialogBtn');
  if (openDialogBtn) {
    openDialogBtn.addEventListener('click', openDialog);
  }
  updateLauncherStatus('Ready to launch', 'info');
}

function openDialog() {
  updateLauncherStatus('Opening dialog...', 'info');

  // Get the dialog URL (same origin as the add-in)
  const dialogUrl = window.location.origin + '/dialog.html';

  console.log('Opening dialog at:', dialogUrl);

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 70,  // 70% of screen height
      width: 60,   // 60% of screen width
      displayInIframe: false  // Open in separate window
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to open dialog:', result.error.message);
        updateLauncherStatus('Failed to open: ' + result.error.message, 'error');
        return;
      }

      dialog = result.value;
      updateLauncherStatus('Dialog opened', 'success');

      // Handle messages from the dialog
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, handleDialogMessage);

      // Handle dialog closed
      dialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent);
    }
  );
}

function handleDialogMessage(arg) {
  console.log('Message from dialog:', arg.message);

  try {
    const message = JSON.parse(arg.message);

    switch (message.type) {
      case 'close':
        console.log('Dialog requested close');
        if (dialog) {
          dialog.close();
          dialog = null;
        }
        updateLauncherStatus('Dialog closed', 'info');
        break;

      case 'generate':
        console.log('Generate request:', message.content);
        updateLauncherStatus('Processing: ' + message.content.substring(0, 30) + '...', 'info');
        // Here we would normally call the WebSocket and insert slides
        // For now, just log it
        handleGenerateRequest(message);
        break;

      default:
        console.log('Unknown message type:', message.type);
    }
  } catch (error) {
    console.error('Failed to parse dialog message:', error);
  }
}

function handleDialogEvent(arg) {
  console.log('Dialog event:', arg);

  switch (arg.error) {
    case 12002: // Dialog closed by user clicking X
      console.log('Dialog closed by user');
      dialog = null;
      updateLauncherStatus('Dialog closed', 'info');
      break;

    case 12003: // Dialog navigated to different domain
      console.log('Dialog navigation error');
      updateLauncherStatus('Navigation error', 'error');
      break;

    case 12006: // Dialog closed programmatically
      console.log('Dialog closed programmatically');
      dialog = null;
      updateLauncherStatus('Ready', 'info');
      break;

    default:
      console.log('Unknown dialog event:', arg.error);
      updateLauncherStatus('Dialog event: ' + arg.error, 'info');
  }
}

function handleGenerateRequest(message) {
  // This is where we would:
  // 1. Send to WebSocket backend
  // 2. Get response
  // 3. Show preview in dialog (send message back)
  // 4. On user confirmation, insert slides

  // For now, just demonstrate with a test slide
  console.log('Would generate content for:', message.content);
  updateLauncherStatus('Received: ' + message.content.substring(0, 40), 'success');

  // TODO: Connect to WebSocket and handle response
  // For demo, we'll just insert a test slide
  insertTestSlide(message.content);
}

async function insertTestSlide(userMessage) {
  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      presentation.slides.load('items');
      await context.sync();

      presentation.slides.add();
      await context.sync();

      presentation.slides.load('items');
      await context.sync();

      const slide = presentation.slides.items[presentation.slides.items.length - 1];

      slide.load('shapes');
      await context.sync();

      // Delete default shapes
      const shapesToDelete = slide.shapes.items.slice();
      for (let shape of shapesToDelete) {
        shape.delete();
      }
      await context.sync();

      // Add title
      const titleShape = slide.shapes.addTextBox('Dialog Test');
      titleShape.left = 50;
      titleShape.top = 50;
      titleShape.width = 600;
      titleShape.height = 60;
      await context.sync();

      titleShape.textFrame.textRange.font.bold = true;
      titleShape.textFrame.textRange.font.size = 32;
      titleShape.textFrame.textRange.font.color = '#d13438';
      await context.sync();

      // Add user message
      const contentShape = slide.shapes.addTextBox('User requested: ' + userMessage);
      contentShape.left = 50;
      contentShape.top = 130;
      contentShape.width = 600;
      contentShape.height = 200;
      await context.sync();

      contentShape.textFrame.textRange.font.size = 18;
      await context.sync();

      console.log('Test slide created');
      updateLauncherStatus('Slide created!', 'success');
    });
  } catch (error) {
    console.error('Error creating slide:', error);
    updateLauncherStatus('Error: ' + error.message, 'error');
  }
}

function updateLauncherStatus(text, type = 'info') {
  const statusEl = document.getElementById('launcherStatus');
  const textEl = document.getElementById('statusText');

  if (textEl) {
    textEl.textContent = text;
  }

  if (statusEl) {
    statusEl.classList.remove('success', 'error');
    if (type === 'success') {
      statusEl.classList.add('success');
    } else if (type === 'error') {
      statusEl.classList.add('error');
    }
  }
}

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

    // Insert into slide
    insertResponseIntoSlide(response);

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
// POWERPOINT SLIDE INSERTION
// ============================================================================

function insertResponseIntoSlide(response) {
  // Don't insert error messages into slides
  if (response.error) {
    console.log('Error response, not inserting into slide');
    return;
  }

  // Check if this is a vocabulary response
  if (response.title && response.words && Array.isArray(response.words)) {
    insertVocabularyIntoSlide(response);
    return;
  }

  // For other content types, use generic insertion
  insertGenericContentIntoSlide(response);
}

async function insertVocabularyIntoSlide(response) {
  try {
    console.log('Inserting vocabulary into slide...', response);

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      // Load slides and add new one
      presentation.slides.load('items');
      await context.sync();

      // slides.add() returns void, not a slide object!
      presentation.slides.add();
      await context.sync();

      // Reload to get the new slide
      presentation.slides.load('items');
      await context.sync();

      // Get the last slide (the one we just added)
      const slide = presentation.slides.items[presentation.slides.items.length - 1];

      // Load shapes collection
      slide.load('shapes');
      await context.sync();

      // Delete default placeholder shapes ("Click to add title", etc.)
      const shapesToDelete = slide.shapes.items.slice();
      for (let shape of shapesToDelete) {
        shape.delete();
      }
      await context.sync();

      // Title text box
      const titleShape = slide.shapes.addTextBox(response.title || 'Vocabulary');
      titleShape.left = 50;
      titleShape.top = 30;
      titleShape.width = 600;
      titleShape.height = 60;
      await context.sync();

      titleShape.textFrame.textRange.font.bold = true;
      titleShape.textFrame.textRange.font.size = 32;
      titleShape.textFrame.textRange.font.color = '#2c3e50';
      await context.sync();

      // Subtitle text box
      if (response.subtitle) {
        const subtitleShape = slide.shapes.addTextBox(response.subtitle);
        subtitleShape.left = 50;
        subtitleShape.top = 95;
        subtitleShape.width = 600;
        subtitleShape.height = 30;
        await context.sync();

        subtitleShape.textFrame.textRange.font.size = 16;
        subtitleShape.textFrame.textRange.font.color = '#7f8c8d';
        await context.sync();
      }

      // Add vocabulary words
      let yPosition = 140;
      const lineHeight = 70;

      for (let index = 0; index < response.words.length; index++) {
        const word = response.words[index];

        // Word and translation
        const wordText = `${index + 1}. ${word.word} — ${word.translation}`;
        const wordShape = slide.shapes.addTextBox(wordText);
        wordShape.left = 50;
        wordShape.top = yPosition;
        wordShape.width = 600;
        wordShape.height = 30;
        await context.sync();

        wordShape.textFrame.textRange.font.size = 18;
        wordShape.textFrame.textRange.font.bold = true;
        wordShape.textFrame.textRange.font.color = '#34495e';
        await context.sync();

        // Definition
        if (word.definition) {
          const defShape = slide.shapes.addTextBox(word.definition);
          defShape.left = 70;
          defShape.top = yPosition + 30;
          defShape.width = 580;
          defShape.height = 35;
          await context.sync();

          defShape.textFrame.textRange.font.size = 14;
          defShape.textFrame.textRange.font.color = '#5d6d7e';
          await context.sync();
        }

        yPosition += lineHeight;
      }

      await context.sync();
      console.log('Vocabulary slide created successfully');
      addMessage('✓ Content added to new slide', 'ai');
    });

  } catch (error) {
    console.error('Error inserting vocabulary into slide:', error);
    addMessage('Error: Could not insert content into slide. ' + error.message, 'ai');
  }
}

async function insertGenericContentIntoSlide(response) {
  try {
    console.log('Inserting generic content into slide...', response);

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      // Load slides and add new one
      presentation.slides.load('items');
      await context.sync();

      // slides.add() returns void
      presentation.slides.add();
      await context.sync();

      // Reload to get the new slide
      presentation.slides.load('items');
      await context.sync();

      // Get the last slide
      const slide = presentation.slides.items[presentation.slides.items.length - 1];

      // Load shapes
      slide.load('shapes');
      await context.sync();

      // Delete default placeholder shapes
      const shapesToDelete = slide.shapes.items.slice();
      for (let shape of shapesToDelete) {
        shape.delete();
      }
      await context.sync();

      // Title
      const title = response.title || 'Generated Content';
      const titleShape = slide.shapes.addTextBox(title);
      titleShape.left = 50;
      titleShape.top = 30;
      titleShape.width = 600;
      titleShape.height = 60;
      await context.sync();

      titleShape.textFrame.textRange.font.bold = true;
      titleShape.textFrame.textRange.font.size = 28;
      await context.sync();

      // Content (as formatted JSON for now)
      const contentText = JSON.stringify(response, null, 2);
      const contentShape = slide.shapes.addTextBox(contentText);
      contentShape.left = 50;
      contentShape.top = 100;
      contentShape.width = 600;
      contentShape.height = 400;
      await context.sync();

      contentShape.textFrame.textRange.font.size = 12;
      contentShape.textFrame.textRange.font.name = 'Courier New';
      await context.sync();

      console.log('Generic content slide created successfully');
      addMessage('✓ Content added to new slide', 'ai');
    });

  } catch (error) {
    console.error('Error inserting generic content into slide:', error);
    addMessage('Error: Could not insert content into slide. ' + error.message, 'ai');
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

  // Build WebSocket message payload
  const messagePayload = {
    "user-id": USER_ID,
    "channel-name": CHANNEL_NAME,
    "conversation-id": null, // null for new conversations (will track later)
    "type": currentCategory, // vocabulary, grammar, reading, etc.
    "content": messageText,
    "requirements": {}
  };

  console.log('Sending message with payload:', messagePayload);

  // Add loading indicator
  addLoadingMessage();

  // Send via WebSocket
  const sent = sendWebSocketMessage(messagePayload);

  // If send failed, remove loading indicator
  if (!sent) {
    removeLoadingMessage();
  }
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
// REFERENCE: Old working slide insertion code from previous version
// This shows the correct pattern for PowerPoint API usage
// Key insights:
// 1. slides.add() returns void, not a slide object
// 2. Must reload slides collection after adding
// 3. Get slide from slides.items[slides.items.length - 1]
// 4. Delete default placeholder shapes before adding custom content
// 5. Sync frequently after operations

async function insertSlidesFromResult(result) {
  if (typeof PowerPoint === 'undefined') {
    throw new Error('PowerPoint API is not available. Please make sure you are running in PowerPoint Desktop.');
  }

  if (!Office.context.requirements || !Office.context.requirements.isSetSupported) {
    throw new Error('API requirements check not available. Please use a newer version of PowerPoint.');
  }

  const v13 = Office.context.requirements.isSetSupported('PowerPointApi', '1.3');
  const isOnline = Office.context.diagnostics.platform === 'OfficeOnline';

  if (!v13) {
    return insertTextFallback(result);
  }

  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      // STEP 1: Load existing slides
      presentation.slides.load('items');
      await context.sync();

      const titleSlideData = result.slides.find(s => s.type === 'title');

      if (titleSlideData) {
        const slides = presentation.slides;

        if (typeof slides.add !== 'function') {
          if (isOnline) {
            return insertTextFallback(result);
          }
          throw new Error('slides.add() method not available. Please update to Office 2019 or Microsoft 365.');
        }

        // STEP 2: Add new slide (returns void, not a slide!)
        slides.add();
        await context.sync();

        // STEP 3: Reload slides to get the new one
        presentation.slides.load('items');
        await context.sync();

        // STEP 4: Get the last slide (the one we just added)
        const titleSlide = presentation.slides.items[presentation.slides.items.length - 1];

        // STEP 5: Load shapes collection
        titleSlide.load('shapes');
        await context.sync();

        // STEP 6: Delete default placeholder shapes
        const shapesToDelete = titleSlide.shapes.items.slice();
        for (let shape of shapesToDelete) {
          shape.delete();
        }
        await context.sync();

        // STEP 7: Add custom title text box
        const title = titleSlide.shapes.addTextBox(titleSlideData.title, {
          left: 50,
          top: 200,
          width: 620,
          height: 100
        });
        await context.sync();

        title.textFrame.textRange.font.size = 44;
        title.textFrame.textRange.font.bold = true;
        title.textFrame.textRange.font.color = "#d13438";
        await context.sync();

        if (titleSlideData.subtitle) {
          const subtitle = titleSlide.shapes.addTextBox(titleSlideData.subtitle, {
            left: 50,
            top: 320,
            width: 620,
            height: 60
          });
          await context.sync();

          subtitle.textFrame.textRange.font.size = 20;
          subtitle.textFrame.textRange.font.color = "#323130";
          await context.sync();
        }
      }

      // Add content slides (vocabulary words)
      const contentData = result.slides.find(s => s.type === 'content');
      if (contentData && contentData.content) {
        for (let i = 0; i < contentData.content.length; i++) {
          const word = contentData.content[i];

          presentation.slides.add();
          await context.sync();

          presentation.slides.load('items');
          await context.sync();

          const wordSlide = presentation.slides.items[presentation.slides.items.length - 1];

          wordSlide.load('shapes');
          await context.sync();

          const shapesToDelete = wordSlide.shapes.items.slice();
          for (let shape of shapesToDelete) {
            shape.delete();
          }
          await context.sync();

          const wordTitle = wordSlide.shapes.addTextBox(word.word, {
            left: 50,
            top: 40,
            width: 620,
            height: 80
          });
          await context.sync();

          wordTitle.textFrame.textRange.font.size = 40;
          wordTitle.textFrame.textRange.font.bold = true;
          wordTitle.textFrame.textRange.font.color = "#d13438";
          await context.sync();

          let bodyText = `Definition: ${word.definition}`;
          if (word.translation) {
            bodyText += `\n\nTranslation: ${word.translation}`;
          }
          if (word.example) {
            bodyText += `\n\nExample: "${word.example}"`;
          }

          const body = wordSlide.shapes.addTextBox(bodyText, {
            left: 50,
            top: 140,
            width: 620,
            height: 360
          });
          await context.sync();

          body.textFrame.textRange.font.size = 18;
          body.textFrame.textRange.font.color = "#323130";
          await context.sync();
        }
      }

      await context.sync();
    });
  } catch (error) {
    if (isOnline) {
      try {
        return await insertTextFallback(result);
      } catch (fallbackError) {
        throw new Error(`Failed to create slides and text fallback also failed: ${fallbackError.message}`);
      }
    }

    throw new Error(`Failed to create slides: ${error.message}`);
  }
}
*/
