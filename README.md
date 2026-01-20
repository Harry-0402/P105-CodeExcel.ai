# CodeExcel AI Assistant

A production-ready Excel Add-in that integrates AI capabilities using the OpenRouter API and Google's Gemini 2.0 Flash model.

## Features

✨ **Two-Tab Interface**
- **Assistant Tab**: Process Excel data with AI
- **Settings Tab**: Secure API key management

🔒 **Secure Credential Handling**
- API keys stored in browser localStorage, not hardcoded
- No credentials in source code or repository

🤖 **AI Integration**
- Uses Google Gemini 2.0 Flash model (free tier available)
- OpenRouter API with required HTTP headers
- System prompt customization

📊 **Excel Integration**
- Reads current cell selection
- Writes AI responses to adjacent cell
- Status indicators and error handling

## Project Structure

```
├── manifest.xml                 # Office Add-in configuration
├── package.json                # Node.js dependencies
├── .gitignore                  # Git ignore rules
├── README.md                   # This file
└── src/
    └── taskpane/
        ├── html/
        │   └── taskpane.html   # UI structure
        ├── css/
        │   └── taskpane.css    # Styling (professional dark theme)
        └── js/
            ├── storage.js      # localStorage wrapper for secure key storage
            ├── api.js          # OpenRouter API integration
            ├── ui.js           # UI state management and status indicators
            └── taskpane.js     # Main app logic and event handlers
```

## Getting Started

### Prerequisites

- **Node.js 14+** (for development server)
- **Excel Online** or **Excel for Web** (Office 365)
- **OpenRouter API Key** ([get free key here](https://openrouter.ai))
- **Modern browser** with localStorage support

### Installation

1. **Clone or download the project**
   ```bash
   cd "P105 CodeExcel.Ai"
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start development server**
   ```bash
   npm start
   ```
   Server runs on `http://localhost:3000`

## Configuration

### 1. Get an OpenRouter API Key

1. Visit [OpenRouter.ai](https://openrouter.ai)
2. Sign up (free account available)
3. Create an API key
4. Copy the key

### 2. Update Manifest

Edit `manifest.xml` and replace the placeholder URL with your hosting URL:

```xml
<SourceLocation DefaultValue="https://your-domain.com/taskpane.html"/>
```

For **local development**, you may need to use a tool like ngrok to expose your local server:
```bash
ngrok http 3000
```

Then update manifest.xml with: `https://your-ngrok-url.ngrok.io/taskpane.html`

### 3. Upload to Excel

1. Open **Excel on the Web** (office.com)
2. Go to **Insert** > **Office Add-ins** > **My Add-ins** > **Upload My Add-in**
3. Select the `manifest.xml` file
4. Click **Upload**

### 4. Configure API Key in Add-in

1. Open the Excel Add-in in the task pane
2. Click the **Settings** tab
3. Paste your OpenRouter API key
4. Optionally customize the AI model and system prompt
5. Click **Save Settings**

The API key is now stored securely in your browser's localStorage and will persist across sessions.

## Usage

1. **Select a cell** in Excel containing data you want to analyze
2. **Click "Process Selection"** in the Assistant tab
3. The AI will analyze the content and write the response to the **adjacent cell (one column right)**
4. **Status indicator** shows: Processing → Success or Error

## API Integration Details

### Headers Required

The API module automatically includes these required headers:

```javascript
{
    "Authorization": `Bearer ${apiKey}`,
    "HTTP-Referer": "https://codeexcel.ai",
    "X-Title": "CodeExcel AI Assistant",
    "Content-Type": "application/json"
}
```

### Model Configuration

Default model: `google/gemini-2.0-flash-lite-preview-02-05:free`

Customizable in Settings tab.

### System Prompt

Default: "You are an Excel data analyst. Complete the following task or analyze this data concisely."

Customizable in Settings tab.

## Code Quality

✅ **Modular JavaScript**
- Separation of concerns (storage, API, UI, main logic)
- No global state pollution
- Reusable components

✅ **Error Handling**
- Network error detection
- API error messages
- Missing credential warnings
- Status feedback to users

✅ **Security**
- No hardcoded API keys
- localStorage for client-side persistence
- Proper CORS headers for OpenRouter

## Deployment to GitHub Pages

### 1. Initialize Git Repository

```bash
git init
git add .
git commit -m "Initial commit: CodeExcel AI Assistant"
```

### 2. Create GitHub Repository

1. Go to GitHub and create a new repository
2. Copy the repository URL

### 3. Connect and Push

```bash
git remote add origin https://github.com/yourusername/codeexcel-ai.git
git branch -M main
git push -u origin main
```

### 4. Enable GitHub Pages

1. Go to repository **Settings** > **Pages**
2. Select **Deploy from a branch**
3. Choose **main branch** and **/ (root)** folder
4. Click **Save**

### 5. Update Manifest

Update `manifest.xml` with your GitHub Pages URL:

```xml
<SourceLocation DefaultValue="https://yourusername.github.io/codeexcel-ai/src/taskpane/html/taskpane.html"/>
```

Push the updated manifest:
```bash
git add manifest.xml
git commit -m "Update manifest with GitHub Pages URL"
git push
```

## Testing Checklist

- [ ] API key saved and persists across sessions
- [ ] "Process Selection" button disabled when no API key
- [ ] Status indicators show correct states
- [ ] AI response displays in response box
- [ ] Response writes to adjacent cell
- [ ] Error messages display for network issues
- [ ] Error messages display for empty cells
- [ ] Settings can be updated
- [ ] API key can be cleared

## Troubleshooting

### "Missing API Key" Error
- Go to Settings tab
- Paste your OpenRouter API key
- Click Save Settings

### "Network Error"
- Check internet connection
- Verify API key is valid
- Check browser console for detailed errors (F12)
- Ensure manifest URL is accessible

### Excel Integration Not Working
- Ensure you're using Excel Online or Excel for Web
- Check that manifest.xml has correct task pane URL
- Verify add-in is properly uploaded

### localStorage Not Persisting
- Check browser's privacy settings
- Some browsers block localStorage in private mode
- Ensure localhost (if using ngrok) is allowed

## Performance Notes

- AI response time depends on OpenRouter API latency
- Typical response: 2-5 seconds
- Status indicator shows processing state
- Max tokens limited to 500 to ensure quick responses

## Browser Compatibility

- ✅ Chrome/Edge (Chromium)
- ✅ Firefox
- ✅ Safari
- ⚠️ IE 11 (not recommended)

## Security Best Practices

1. **API Key**: Never commit to version control (covered by .gitignore)
2. **HTTPS**: Always use HTTPS in production
3. **Validation**: User input is escaped before display
4. **CORS**: OpenRouter handles CORS with proper headers

## Development Tips

### Hot Reload
For development with auto-reload:
```bash
npm run dev
```

### Console Debugging
All modules log to console. Open browser DevTools (F12) to see:
- API calls
- Storage operations
- Excel integration events

### Common Issues

**localStorage blocked?**
- Check browser privacy/incognito mode
- Ensure cookies are allowed for the domain

**Excel Add-in won't load?**
- Check manifest.xml SourceLocation URL is accessible
- Verify HTTPS (if applicable)
- Check browser console for errors

## Future Enhancements

- [ ] Support for multiple cell ranges
- [ ] Batch processing
- [ ] Custom system prompts per task
- [ ] Response history
- [ ] Model selection UI
- [ ] Rate limiting and quota management
- [ ] User authentication
- [ ] Database integration for storing responses

## License

MIT License - See LICENSE file

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review browser console errors (F12)
3. Verify manifest.xml configuration
4. Test API key with OpenRouter directly

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

---

**Created**: January 2026
**Status**: Production Ready
**Office.js Version**: 1.1.97+
**Node.js Version**: 14+
