/* global document, Office, PowerPoint */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    initializeUI();
  }
});

function initializeUI() {
  document.getElementById("generateBtn").onclick = generateVocabulary;
  document.getElementById("generateAdvancedBtn").onclick = generateVocabulary;
  document.getElementById("advancedBtn").onclick = toggleAdvancedOptions;

  const slider = document.getElementById('vocabWordCount');
  const display = document.getElementById('wordCountValue');

  if (slider && display) {
    slider.addEventListener('input', function() {
      display.textContent = this.value;
      updateSmartDefaultsDisplay();
    });
  }

  const inputs = ['vocabLanguage', 'vocabExamples', 'vocabImages'];
  inputs.forEach(id => {
    const element = document.getElementById(id);
    if (element) {
      element.addEventListener('change', updateSmartDefaultsDisplay);
    }
  });

  updateSmartDefaultsDisplay();
}

function updateSmartDefaultsDisplay() {
  const language = document.getElementById('vocabLanguage').value;
  const wordCount = document.getElementById('vocabWordCount').value;
  const examples = document.getElementById('vocabExamples').checked;

  const languageNames = {
    'en': 'English',
    'de': 'German',
    'fr': 'French',
    'es': 'Spanish',
    'it': 'Italian',
    'pt': 'Portuguese'
  };

  const langName = languageNames[language] || 'English';
  const exampleText = examples ? 'with examples' : 'no examples';

  const smartText = document.querySelector('.smart-text');
  if (smartText) {
    smartText.textContent = `Auto-selected: ${langName}, ${wordCount} words, ${exampleText}`;
  }
}

function toggleAdvancedOptions() {
  const advancedOptions = document.getElementById('advancedOptions');
  const icon = document.getElementById('advancedIcon');
  const buttonText = document.getElementById('advancedBtn').querySelector('span:last-child');

  if (advancedOptions.style.display === 'none') {
    advancedOptions.style.display = 'block';
    icon.textContent = 'expand_less';
    buttonText.textContent = 'Fewer Options';
  } else {
    advancedOptions.style.display = 'none';
    icon.textContent = 'expand_more';
    buttonText.textContent = 'More Options';
  }
}

function showStatus(message, isSuccess = true) {
  const status = document.getElementById('status');
  status.textContent = message;
  status.className = `status ${isSuccess ? 'success' : 'error'}`;
  status.style.display = 'block';

  setTimeout(() => {
    status.style.display = 'none';
  }, 5000);
}


function validateForm() {
  const topic = document.getElementById('vocabTopic').value.trim();
  const level = document.getElementById('vocabLevel').value;

  const errors = [];

  if (!topic) {
    errors.push('Topic is required');
  }

  if (topic && topic.length < 3) {
    errors.push('Please enter a more specific topic (3+ characters)');
  }

  if (!level) {
    errors.push('Student level is required');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

function getFormData() {
  return {
    content_type: 'vocabulary',
    language: document.getElementById('vocabLanguage').value,
    level: document.getElementById('vocabLevel').value,
    parameters: {
      topic: document.getElementById('vocabTopic').value.trim(),
      word_count: parseInt(document.getElementById('vocabWordCount').value),
      include_examples: document.getElementById('vocabExamples').checked,
      include_images: document.getElementById('vocabImages').checked
    }
  };
}

async function generateVocabulary() {
  const validation = validateForm();
  if (!validation.isValid) {
    showStatus(validation.errors.join('. '), false);
    return;
  }

  const formData = getFormData();

  const button = event.target;
  button.disabled = true;
  const originalHTML = button.innerHTML;

  button.innerHTML = '<span class="material-icons button-icon">hourglass_empty</span>Generating...';

  const topic = document.getElementById('vocabTopic').value;
  const level = document.getElementById('vocabLevel').value;
  showStatus(`Creating ${level} vocabulary slides for "${topic}"... This may take 20-30 seconds.`, true);

  try {
    const backendUrl = 'http://localhost:2000/api/generate';

    const response = await fetch(backendUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(formData)
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.status}`);
    }

    const result = await response.json();

    if (!result.success) {
      throw new Error(result.error || 'Failed to generate vocabulary');
    }

    await insertSlidesFromResult(result);

    showStatus(`Successfully created ${result.metadata.word_count} vocabulary slides!`, true);

  } catch (error) {
    console.error('Full error details:', error);
    const errorMessage = error.message || error.toString();
    showStatus(errorMessage, false);
  } finally {
    button.disabled = false;
    button.innerHTML = originalHTML;
  }
}

function insertTextFallback(result) {
  const options = { coercionType: Office.CoercionType.Text };

  let displayText = `${result.slides[0].title}\n`;
  displayText += `${result.slides[0].subtitle || ''}\n\n`;
  displayText += `=== Vocabulary Words ===\n\n`;

  const contentSlide = result.slides.find(slide => slide.type === 'content');
  if (contentSlide && contentSlide.content) {
    contentSlide.content.forEach((word, index) => {
      displayText += `${index + 1}. ${word.word.toUpperCase()}\n\n`;
      displayText += `   Definition: ${word.definition}\n\n`;
      if (word.translation) {
        displayText += `   Translation: ${word.translation}\n\n`;
      }
      if (word.example) {
        displayText += `   Example: ${word.example}\n\n`;
      }
      displayText += `${'='.repeat(50)}\n\n`;
    });
  }

  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(displayText, options, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(asyncResult.error.message));
      }
    });
  });
}

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

        // slides.add() returns void, not a slide object
        slides.add();
        await context.sync();

        presentation.slides.load('items');
        await context.sync();

        const titleSlide = presentation.slides.items[presentation.slides.items.length - 1];

        titleSlide.load('shapes');
        await context.sync();

        const shapesToDelete = titleSlide.shapes.items.slice();
        for (let shape of shapesToDelete) {
          shape.delete();
        }
        await context.sync();

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

export { generateVocabulary };
