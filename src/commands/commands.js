/* global Office, PowerPoint */

/**
 * Teachers Center - Commands Module
 * Handles dialog lifecycle and PowerPoint slide operations
 */

let dialog = null;

Office.onReady(() => {
    console.log('Commands.js: Office is ready');
});

// ============================================
// DIALOG MANAGEMENT
// ============================================

/**
 * Opens the Teachers Center dialog directly from the ribbon button
 * @param {Office.AddinCommands.Event} event
 */
function openTeachersCenterDialog(event) {
    console.log('Opening Teachers Center dialog...');

    // If dialog is already open, just complete the event
    if (dialog) {
        console.log('Dialog already open');
        event.completed();
        return;
    }

    const dialogUrl = `${window.location.origin}/dialog.html`;
    console.log('Dialog URL:', dialogUrl);

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        {
            height: 70,  // 70% of screen height
            width: 60,   // 60% of screen width
            displayInIframe: false
        },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to open dialog:', result.error.code, result.error.message);
                event.completed();
                return;
            }

            dialog = result.value;
            console.log('Dialog opened successfully');

            // Handle messages from the dialog
            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                (arg) => handleDialogMessage(arg, event)
            );

            // Handle dialog events (closed, etc.)
            dialog.addEventHandler(
                Office.EventType.DialogEventReceived,
                (arg) => handleDialogEvent(arg, event)
            );

            // Complete the event - dialog is now open
            event.completed();
        }
    );
}

/**
 * Handle messages received from the dialog
 */
function handleDialogMessage(arg, event) {
    console.log('Message from dialog:', arg.message);

    try {
        const message = JSON.parse(arg.message);

        switch (message.type) {
            case 'close':
                closeDialog();
                break;

            case 'generate':
                handleGenerateRequest(message);
                break;

            case 'insert':
                handleInsertRequest(message);
                break;

            case 'cancel':
                // Handle cancellation if needed
                console.log('Operation cancelled by user');
                break;

            default:
                console.log('Unknown message type:', message.type);
        }
    } catch (error) {
        console.error('Failed to parse dialog message:', error);
    }
}

/**
 * Handle dialog events (close, navigation errors, etc.)
 */
function handleDialogEvent(arg, event) {
    console.log('Dialog event:', arg.error);

    switch (arg.error) {
        case 12002: // Dialog closed by user clicking X
        case 12006: // Dialog closed programmatically
            console.log('Dialog closed');
            dialog = null;
            break;

        case 12003: // Dialog navigated to different domain
            console.error('Dialog navigation error');
            dialog = null;
            break;

        default:
            console.log('Unknown dialog event:', arg.error);
    }
}

/**
 * Close the dialog
 */
function closeDialog() {
    if (dialog) {
        dialog.close();
        dialog = null;
    }
}

/**
 * Send a message to the dialog
 */
function sendToDialog(message) {
    if (dialog) {
        try {
            dialog.messageChild(JSON.stringify(message));
        } catch (error) {
            console.error('Failed to send message to dialog:', error);
        }
    }
}

// ============================================
// CONTENT GENERATION
// ============================================

/**
 * Handle generate content request from dialog
 * This will be connected to a backend/AI service in production
 */
async function handleGenerateRequest(message) {
    console.log('Generate request:', message);

    // Send progress updates to dialog
    sendToDialog({ type: 'progress', stage: 'Analyzing request...', percent: 10 });

    // Simulate AI processing with progress updates
    await simulateProgress([
        { stage: 'Analyzing request...', percent: 20, delay: 300 },
        { stage: 'Generating vocabulary...', percent: 40, delay: 500 },
        { stage: 'Creating slide content...', percent: 60, delay: 400 },
        { stage: 'Formatting slides...', percent: 80, delay: 300 },
        { stage: 'Preparing preview...', percent: 95, delay: 200 }
    ]);

    // Generate demo slides (replace with actual AI response in production)
    const slides = generateDemoSlides(message.content, message.category);

    // Send preview to dialog
    sendToDialog({
        type: 'preview',
        slides: slides,
        summary: `Generated ${slides.length} slides`
    });
}

/**
 * Simulate progress updates with delays
 */
async function simulateProgress(stages) {
    for (const stage of stages) {
        sendToDialog({ type: 'progress', stage: stage.stage, percent: stage.percent });
        await sleep(stage.delay);
    }
}

/**
 * Generate demo slides based on content and category
 * Replace with actual AI-generated content in production
 */
function generateDemoSlides(content, category) {
    // Extract topic from content (simple parsing)
    const topic = content.toLowerCase().includes('german')
        ? 'German Food Vocabulary'
        : content.substring(0, 50);

    const slides = [
        {
            type: 'Title',
            title: topic,
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

// ============================================
// SLIDE INSERTION
// ============================================

/**
 * Handle insert slides request from dialog
 */
async function handleInsertRequest(message) {
    console.log('Insert request:', message);

    const slides = message.slides || [];
    const total = slides.length;

    if (total === 0) {
        sendToDialog({
            type: 'error',
            message: 'No slides to insert.'
        });
        return;
    }

    try {
        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;

            for (let i = 0; i < slides.length; i++) {
                const slideData = slides[i];

                // Send progress update
                sendToDialog({
                    type: 'insertProgress',
                    current: i + 1,
                    total: total
                });

                // Add new slide
                presentation.slides.add();
                await context.sync();

                // Load slides to get the new one
                presentation.slides.load('items');
                await context.sync();

                const slide = presentation.slides.items[presentation.slides.items.length - 1];

                // Clear default placeholders
                slide.shapes.load('items');
                await context.sync();

                const shapesToDelete = slide.shapes.items.slice();
                for (const shape of shapesToDelete) {
                    shape.delete();
                }
                await context.sync();

                // Create slide content based on type
                await createSlideContent(slide, slideData, context);

                // Small delay between slides for visual feedback
                await sleep(200);
            }

            // Final sync
            await context.sync();

            // Send success message
            sendToDialog({
                type: 'success',
                message: `${total} slide${total !== 1 ? 's' : ''} inserted successfully`
            });

            console.log(`Successfully inserted ${total} slides`);
        });
    } catch (error) {
        console.error('Error inserting slides:', error);
        sendToDialog({
            type: 'error',
            message: `Failed to insert slides: ${error.message}`
        });
    }
}

/**
 * Create content for a single slide
 */
async function createSlideContent(slide, slideData, context) {
    const isTitle = slideData.type === 'Title';

    // Title text box
    const titleShape = slide.shapes.addTextBox(slideData.title || '');
    titleShape.left = 50;
    titleShape.top = isTitle ? 180 : 40;
    titleShape.width = 620;
    titleShape.height = isTitle ? 80 : 60;
    await context.sync();

    titleShape.textFrame.textRange.font.bold = true;
    titleShape.textFrame.textRange.font.size = isTitle ? 44 : 32;
    titleShape.textFrame.textRange.font.color = '#d13438';
    if (isTitle) {
        titleShape.textFrame.horizontalAlignment = 'Center';
    }
    await context.sync();

    // Subtitle text box
    if (slideData.subtitle) {
        const subtitleShape = slide.shapes.addTextBox(slideData.subtitle);
        subtitleShape.left = 50;
        subtitleShape.top = isTitle ? 270 : 100;
        subtitleShape.width = 620;
        subtitleShape.height = 40;
        await context.sync();

        subtitleShape.textFrame.textRange.font.size = isTitle ? 24 : 20;
        subtitleShape.textFrame.textRange.font.color = '#605e5c';
        if (isTitle) {
            subtitleShape.textFrame.horizontalAlignment = 'Center';
        }
        await context.sync();
    }

    // Content text box (for vocabulary cards)
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

    // Example sentence (for vocabulary cards)
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
// UTILITY FUNCTIONS
// ============================================

/**
 * Sleep utility for delays
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================
// OFFICE ACTIONS REGISTRATION
// ============================================

// Register the function with Office
Office.actions.associate("openTeachersCenterDialog", openTeachersCenterDialog);
