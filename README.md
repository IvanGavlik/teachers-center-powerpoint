# Teachers Center PowerPoint Add-in

A PowerPoint add-in that generates vocabulary slides for language teachers. Creates interactive, AI-powered vocabulary presentations with definitions, translations, and examples.

## Features

- Generate vocabulary slides based on topic and student level
- Customizable word count (5-15 words)
- Optional examples and translations
- Multiple language support (English, German, French, Spanish, Italian, Portuguese)
- Automatic slide formatting with clean design

## Prerequisites

- Node.js (v14 or higher)
- npm (v6 or higher)
- PowerPoint Desktop (Office 2016 or later, or Microsoft 365)
- Backend API running on `http://localhost:2000`

## Important Limitations

**This add-in does NOT work with PowerPoint Online.** The `slides.add()` API used to create slides programmatically has limited support in PowerPoint Online. Please use PowerPoint Desktop for full functionality.

## Workspace Setup

### 1. Clone the repository

```bash
git clone <repository-url>
cd teachers-center-powerpoint
```

### 2. Install dependencies

```bash
npm install
```

### 3. Generate SSL certificates

Office Add-ins require HTTPS. Generate development certificates:

```bash
npx office-addin-dev-certs install
```

## Running the Application Locally

### 1. Start the development server

```bash
npm run dev-server
```

This will start the webpack dev server on `https://localhost:3000`.

### 2. Start the add-in in PowerPoint

In a separate terminal:

```bash
npm start
```

This command will:
- Open PowerPoint Desktop
- Sideload the add-in automatically
- Display the add-in task pane

### 3. Stop the add-in

```bash
npm stop
```

## Backend API Endpoint

The add-in communicates with a backend API to generate vocabulary content.

### Endpoint Configuration

- **URL**: `http://localhost:2000/api/generate`
- **Method**: POST
- **Content-Type**: application/json

### Request Format

```json
{
  "content_type": "vocabulary",
  "language": "en",
  "level": "B1",
  "parameters": {
    "topic": "food",
    "word_count": 10,
    "include_examples": true,
    "include_images": false
  }
}
```

### Response Format

```json
{
  "success": true,
  "slides": [
    {
      "type": "title",
      "title": "Vocabulary: Food",
      "subtitle": "B1 Level"
    },
    {
      "type": "content",
      "content": [
        {
          "word": "recipe",
          "definition": "A set of instructions for cooking",
          "translation": "Rezept",
          "example": "Follow the recipe to make the cake."
        }
      ]
    }
  ],
  "metadata": {
    "word_count": 10
  }
}
```

## Testing Locally

### 1. Start the backend server

Make sure your backend API is running on `http://localhost:2000` before testing.

### 2. Launch the add-in

```bash
npm run dev-server
npm start
```

### 3. Test the vocabulary generator

1. In PowerPoint, open the add-in task pane from the Home ribbon
2. Enter a topic (e.g., "food", "travel", "business")
3. Select student level (A1-C2)
4. Adjust word count if needed (5-15 words)
5. Click "Quick Generate"
6. Wait 20-30 seconds for slide generation

### 4. Verify generated slides

The add-in should create:
- 1 title slide with topic and level
- N content slides (one per vocabulary word)
- Each word slide contains: word, definition, translation, and example

## Development Commands

### Build for production

```bash
npm run build
```

### Build for development

```bash
npm run build:dev
```

### Watch mode (auto-rebuild on changes)

```bash
npm watch
```

### Validate manifest

```bash
npm run validate
```

### Lint code

```bash
npm run lint
npm run lint:fix
```

## Project Structure

```
teachers-center-powerpoint/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Main UI
│   │   ├── taskpane.css     # Styles
│   │   └── taskpane.js      # Business logic
│   └── commands/
│       ├── commands.html
│       └── commands.js
├── manifest.xml              # Add-in configuration
├── webpack.config.js         # Build configuration
└── package.json              # Dependencies and scripts
```

## PowerPoint API Requirements

This add-in requires **PowerPointApi 1.3** or higher. The following Office versions are supported:

- Office 2016 or later
- Microsoft 365

### Key API Features Used

- `PowerPoint.run()` - Execute operations in batch
- `presentation.slides.add()` - Create new slides (returns void)
- `slide.shapes.addTextBox()` - Add text content
- `shape.textFrame.textRange.font` - Format text

## Troubleshooting

### Add-in doesn't load

- Ensure PowerPoint Desktop is installed (not Online)
- Clear Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
- Restart PowerPoint completely

### Slides not created

- Verify PowerPointApi 1.3+ support
- Check if using PowerPoint Online (not supported)
- Ensure backend API is running and accessible

### SSL certificate errors

- Reinstall dev certificates: `npx office-addin-dev-certs install`
- Trust the certificate in your system

### Backend connection errors

- Verify backend is running on `http://localhost:2000`
- Check CORS settings on backend
- Review browser console for fetch errors

## License

MIT
