# Quick Start Guide - CodeExcel AI

Get up and running in 5 minutes.

## Step 1: Get Your API Key (2 minutes)

1. Go to https://openrouter.ai
2. Sign up → Generate API Key
3. **Copy your key** (keep it safe!)

## Step 2: Prepare Your Computer (1 minute)

**Option A: With Node.js** (recommended)
```bash
npm install
npm start
```
Server starts at http://localhost:3000

**Option B: Without Node.js**
- Use any HTTP server (Python, PHP, etc.)
- Or use [ngrok](https://ngrok.com) to tunnel local files

## Step 3: Upload to Excel Online (1 minute)

1. Open https://office.com → Excel
2. Click **Insert** → **Office Add-ins** → **Upload My Add-in**
3. Select the `manifest.xml` file in this project
4. Click **Upload**

## Step 4: Configure API Key (30 seconds)

1. In the task pane, click the **Settings** tab
2. Paste your OpenRouter API key
3. Click **Save Settings**
4. ✅ Done! Your key is securely stored in browser

## Step 5: Use It! (30 seconds)

1. Click the **Assistant** tab
2. Select any Excel cell with text
3. Click **Process Selection**
4. Watch as AI analyzes your data
5. Results appear in the next column

---

## Troubleshooting

### "API Key required" error
→ Go to Settings, add your API key, click Save

### Add-in won't load
→ Check that taskpane.html URL is accessible

### Slow responses
→ Normal (2-5 seconds). OpenRouter/Gemini may be busy.

### Key not saving
→ Enable localStorage (not in incognito mode)

---

## What Just Happened?

1. **You uploaded** an Excel Add-in manifest
2. **Excel loaded** the task pane from your server
3. **You saved** your API key to browser storage (NOT your server)
4. **You sent** cell data to OpenRouter's Gemini model
5. **AI processed** your data and returned results
6. **Results wrote** back to Excel automatically

## Next Steps

- Read [README.md](README.md) for full documentation
- Check [OPENROUTER_API.md](OPENROUTER_API.md) to customize AI settings
- Deploy to GitHub Pages using [README.md](README.md) instructions
- Review [DEPLOYMENT_CHECKLIST.md](DEPLOYMENT_CHECKLIST.md) before production

## File Structure Overview

```
manifest.xml ← Excel reads this to find your add-in
src/taskpane/
  html/taskpane.html ← User interface
  css/taskpane.css ← Styling
  js/
    storage.js ← Saves your API key
    api.js ← Talks to OpenRouter
    ui.js ← Updates status messages
    taskpane.js ← Main logic
```

## Key Features

✅ **Secure**: Your API key never leaves your browser
✅ **Free**: Uses free Gemini model on OpenRouter
✅ **Fast**: 2-5 second response times
✅ **Easy**: No coding required after setup
✅ **Customizable**: Change AI model and instructions

---

**Need help?** Check README.md → Troubleshooting section

**Ready to deploy?** See README.md → Deployment to GitHub Pages

Happy analyzing! 🚀
