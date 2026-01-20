# Troubleshooting Guide - CodeExcel AI Assistant

Comprehensive solutions for common issues.

## Connection & API Issues

### ❌ "Error: Missing API Key"

**Cause**: No API key saved or it was cleared

**Solutions**:
1. Go to **Settings** tab in the add-in
2. Paste your OpenRouter API key
3. Click **Save Settings**
4. Return to **Assistant** tab
5. Try processing again

**Verify**:
- Open browser DevTools (F12)
- Go to **Application** → **Local Storage**
- Check if `openrouter_api_key` exists

---

### ❌ "Error: API Error: HTTP 401"

**Cause**: Invalid or expired API key

**Solutions**:
1. Get a new API key from https://openrouter.ai
2. Go to Settings tab
3. Delete old key (click **Clear API Key**)
4. Paste new key
5. Click **Save Settings**

**Verify**:
- Test key at https://openrouter.ai/account/usage (should show your account)

---

### ❌ "Error: API Error: HTTP 429"

**Cause**: Rate limit exceeded (too many requests)

**Solutions**:
1. **Wait 1 minute** before trying again
2. Check your usage at https://openrouter.ai/account/usage
3. If on free tier, you have limited quota
4. Consider upgrading to paid plan for higher limits

**Prevent**:
- Don't send multiple requests rapidly
- Process cells one at a time
- Wait 2-3 seconds between requests

---

### ❌ "Error: Network error: Unable to reach the API"

**Cause**: No internet connection or API unreachable

**Solutions**:
1. **Check internet**: Open any website in browser
2. **Verify OpenRouter is up**: Visit https://status.openrouter.io
3. **Refresh the page**: Ctrl+R (or Cmd+R on Mac)
4. **Check firewall**: Ensure port 443 (HTTPS) is open
5. **Try ngrok if on corporate network**: 
   ```bash
   ngrok http 3000
   ```
   Then update manifest.xml with ngrok URL

**Check Network**:
- Open DevTools (F12) → **Network** tab
- Attempt to process cell
- Look for failed requests to openrouter.ai
- Check response details for error message

---

### ❌ "Error: Invalid API response format"

**Cause**: API returned unexpected data format

**Solutions**:
1. Verify API key is valid (see HTTP 401 solution)
2. Verify model name is correct in Settings
3. Try a different model:
   - Click **Settings**
   - Change model to: `openai/gpt-3.5-turbo`
   - Save and try again
4. Clear browser cache (Ctrl+Shift+Delete)

---

## Excel Integration Issues

### ❌ Add-in doesn't load in Excel

**Cause**: manifest.xml URL is incorrect or unreachable

**Solutions**:

**For Local Development**:
1. Ensure server is running: `npm start`
2. Check it's accessible: Open http://localhost:3000/taskpane.html in browser
3. If on corporate network, use ngrok:
   ```bash
   ngrok http 3000
   ```
4. Update `manifest.xml`:
   ```xml
   <SourceLocation DefaultValue="https://your-ngrok-url.ngrok.io/taskpane.html"/>
   ```

**For GitHub Pages**:
1. Verify repository is public
2. GitHub Pages is enabled (Settings → Pages)
3. `manifest.xml` has correct URL:
   ```xml
   <SourceLocation DefaultValue="https://username.github.io/codeexcel-ai/src/taskpane/html/taskpane.html"/>
   ```
4. File exists at that URL (test in browser)

**Verify Manifest**:
- Open DevTools (F12) → **Console**
- Look for errors loading taskpane.html
- Check Network tab for 404 errors

---

### ❌ "Selected Cell" shows "None"

**Cause**: Cell selection tracking not working

**Solutions**:
1. Click on a cell in Excel (make sure it's selected)
2. The display should update automatically
3. If not:
   - Refresh the page (Ctrl+R)
   - Try selecting a different cell
   - Check browser console for errors

---

### ❌ Response doesn't write to cell

**Cause**: Excel permissions or syntax error

**Solutions**:
1. Verify manifest has `<Permissions>ReadWriteDocument</Permissions>`
2. Check if Excel is in edit mode (try clicking elsewhere first)
3. Try a simpler cell value (just text, no formulas)
4. Check that target cell (adjacent column) is empty
5. Refresh add-in and try again

**Debug**:
- Open DevTools → Console
- Select cell and process
- Look for Excel.run errors
- Target cell should have green background after success

---

## UI & Display Issues

### ❌ Status indicator stuck on "Processing..."

**Cause**: Response timed out or error occurred

**Solutions**:
1. Wait 10 seconds (API may still be processing)
2. Check internet connection
3. Refresh page (Ctrl+R)
4. Try a simpler input (fewer characters)
5. Check DevTools Console for error details

**Prevent**:
- Ensure API key is valid
- Check OpenRouter status
- Reduce input size
- Try during off-peak hours

---

### ❌ Settings don't save

**Cause**: localStorage blocked or quota exceeded

**Solutions**:
1. **Check if not in private mode**: Exit private/incognito mode
2. **Allow localStorage**:
   - Chrome: Settings → Privacy → Cookies → Allow all
   - Firefox: Preferences → Privacy → Cookies → Allow all
3. **Clear storage quota**:
   - Open DevTools (F12) → Application → Clear site data
   - Return to add-in and try again
4. **Check storage size**:
   - DevTools → Application → Local Storage → Check used space

**Verify**:
- Save settings
- Open DevTools → Application → Local Storage
- Should see `openrouter_api_key` in the list

---

### ❌ Response text looks garbled

**Cause**: HTML/special characters not escaped

**Solutions**:
- This shouldn't happen - it's automatically escaped
- If it does, clear cache and refresh:
  ```
  Ctrl+Shift+Delete (or Cmd+Shift+Delete on Mac)
  ```
- Check browser console for JavaScript errors

---

## Performance Issues

### ❌ Response is very slow (>10 seconds)

**Cause**: OpenRouter/Gemini overloaded

**Solutions**:
1. **Wait and retry**: Normal during peak hours (9am-5pm)
2. **Try different time**: Try early morning or late evening
3. **Try different model**: Go to Settings → change model → Save
4. **Check internet**: Run speed test at speedtest.net
5. **Restart browser**: Close all tabs and reopen

**Optimize**:
- Reduce input size (shorter cell values process faster)
- Use model with faster response time
- Process during off-peak hours

---

### ❌ Add-in is laggy or unresponsive

**Cause**: Browser resource limitation

**Solutions**:
1. Close other browser tabs
2. Restart browser
3. Check available RAM (Task Manager on Windows)
4. Try different browser
5. Disable browser extensions (Tools → Extensions → Disable All)

---

## Data Loss & Safety

### ❌ Cell data got overwritten

**Cause**: Response written to wrong column

**Solutions**:
1. Use Ctrl+Z (undo) to restore
2. Ensure only one column shift (one to the right)
3. Check that target cell was empty before processing

**Prevent**:
- Always have an empty column to the right
- Test with sample data first
- Backup important data before using

---

## Browser-Specific Issues

### ❌ Works in Chrome but not in Firefox

**Solution**: 
- Firefox may have different localStorage rules
- Check: Preferences → Privacy → Cookies → Allow
- Clear cache and cookies for the site

### ❌ Works on Desktop but not on Mobile

**Solution**:
- Excel add-ins require Excel on Web, not mobile app
- Use Office.com on mobile browser instead

### ❌ Safari: "Cannot read property 'Office' of undefined"

**Solution**:
- Safari may block Office.js
- Check Security settings
- Try updating Safari to latest version
- Use Chrome/Firefox as workaround

---

## Developer Debugging

### Enable Console Logging

All modules log to browser console. To see details:

1. Open DevTools (F12)
2. Go to **Console** tab
3. Process a cell
4. Look for log messages:
   ```
   "Processing cell A1 with value: ..."
   "API response received"
   "Writing to Excel..."
   ```

### Check Storage

1. Open DevTools (F12)
2. Go to **Application** tab
3. **Local Storage** → Select your domain
4. Should see:
   - `openrouter_api_key`
   - `openrouter_model`
   - `openrouter_system_prompt`

### Monitor Network Requests

1. Open DevTools (F12)
2. Go to **Network** tab
3. Process a cell
4. Look for POST request to `openrouter.ai/api/v1/chat/completions`
5. Click it to see:
   - **Request**: Headers (authorization, referer, etc.)
   - **Response**: JSON with AI response

### Check Excel.run() Errors

1. Open DevTools (F12) → **Console**
2. Process a cell
3. Look for `Excel.run` errors
4. Common causes:
   - Wrong range selection
   - Read-only cells
   - No permissions

---

## When All Else Fails

### Complete Reset

1. **Clear everything**:
   - DevTools → Application → Clear site data
   - Close browser entirely
   - Reopen browser

2. **Reinstall add-in**:
   - Excel: Insert → Office Add-ins → Manage My Add-ins
   - Remove CodeExcel AI Assistant
   - Re-upload manifest.xml

3. **Check files**:
   - Verify all files exist in correct folders
   - Verify manifest.xml is valid XML
   - Verify URLs in manifest are accessible

### Get Help

If you're still stuck:

1. **Open DevTools** (F12) and go to **Console**
2. **Reproduce the error**
3. **Take screenshot** of error message
4. **Check**:
   - README.md Troubleshooting
   - OPENROUTER_API.md
   - This document

5. **Log details**:
   - Error message (exact text)
   - Steps to reproduce
   - Browser and OS
   - What you've already tried

---

## Success Indicators ✅

You'll know everything is working when:

- [ ] "Ready" status shows on startup
- [ ] API key saves without errors
- [ ] Selecting cells updates the display
- [ ] Status changes: Processing → Success
- [ ] Response appears in response box
- [ ] Data appears in adjacent Excel cell
- [ ] Green background shows on result cell
- [ ] Can process multiple cells in sequence

---

## Reporting Issues

If you find a bug:

1. **Reproduce** consistently
2. **Document**:
   - Exact error message
   - Steps to reproduce
   - Browser and OS
   - Cell content/format
3. **Check console** for JavaScript errors
4. **Provide**:
   - manifest.xml (with URL redacted)
   - Browser version
   - Network tab screenshot

---

**Last Updated**: January 2026
**Version**: 1.0
**Tested On**: Chrome, Edge, Firefox, Safari
