# Deployment Checklist for CodeExcel AI Assistant

Use this checklist to ensure your Excel Add-in is properly configured and deployed.

## Pre-Deployment Setup

### Code Review
- [ ] All JavaScript files have no syntax errors
- [ ] manifest.xml is properly formatted
- [ ] All file paths are correct in HTML imports
- [ ] No hardcoded API keys or secrets in any files
- [ ] .gitignore includes sensitive files

### Local Testing
- [ ] npm install completes without errors
- [ ] npm start launches development server
- [ ] taskpane.html loads in browser at http://localhost:3000/taskpane.html
- [ ] All UI elements render correctly
- [ ] Tab switching works
- [ ] Form inputs are functional

## OpenRouter Setup

- [ ] Created OpenRouter account
- [ ] Generated API key
- [ ] Copied API key (stored securely, not in code)
- [ ] Verified API key works with test request
- [ ] Reviewed free tier quotas and limits
- [ ] Bookmarked https://openrouter.ai/account/usage for monitoring

## Office Add-in Configuration

### Manifest.xml
- [ ] Updated `<SourceLocation>` with correct taskpane.html URL
- [ ] Confirmed `<DisplayName>` is appropriate
- [ ] Verified `<Description>` is clear
- [ ] Checked that `<Host Name="Workbook"/>` is present
- [ ] `<Permissions>ReadWriteDocument</Permissions>` is set
- [ ] GUID in `<Id>` is unique

### Web Hosting
- [ ] Taskpane.html URL is publicly accessible (or on ngrok for dev)
- [ ] HTTPS is used in production (HTTP ok for localhost)
- [ ] CORS headers are configured (if applicable)
- [ ] Server is responding with 200 status
- [ ] Files are cached-busted or cache control headers are set

## Excel Upload & Testing

### Add-in Upload
- [ ] Opened Excel Online (office.com)
- [ ] Navigated to Insert > Office Add-ins > Upload My Add-in
- [ ] Selected manifest.xml file
- [ ] Clicked Upload
- [ ] Task pane loaded without errors
- [ ] Header and tabs are visible

### UI Testing
- [ ] Selected cell display updates correctly
- [ ] Status indicator shows "Ready" on initial load
- [ ] Assistant tab is active by default
- [ ] Settings tab shows form with API key input
- [ ] Process button is visible in Assistant tab
- [ ] Save button is visible in Settings tab
- [ ] Clear button is visible in Settings tab

### Settings Tab Testing
- [ ] Can enter API key in input field
- [ ] Save button stores key (check localStorage via DevTools)
- [ ] Settings persist after refresh
- [ ] Model input shows default model name
- [ ] System prompt textarea shows default prompt
- [ ] Can modify model and prompt
- [ ] Clear button removes API key

### Assistant Tab Testing
- [ ] With no API key: Process button disabled, error status shown
- [ ] With API key: Process button enabled, status shows "Ready"
- [ ] Select an empty cell: Process button shows error
- [ ] Select a cell with data (e.g., "test"), click Process:
  - [ ] Status changes to "Processing..."
  - [ ] Wait for response (2-5 seconds)
  - [ ] Response appears in response box
  - [ ] Response is written to adjacent cell
  - [ ] Status changes to "Success"

### Error Handling Testing
- [ ] Delete API key, try to process: Shows "API Key required"
- [ ] Select empty cell, try to process: Shows "Selected cell is empty"
- [ ] Disconnect internet, try to process: Shows network error
- [ ] Invalid API key, try to process: Shows API error
- [ ] All errors have clear messages

## Excel Integration Testing

- [ ] Cell selection tracking works
- [ ] "Selected Cell" display updates when clicking cells
- [ ] Response writes to exactly 1 column to the right
- [ ] Response cell has green background (#E8F5E9)
- [ ] Multi-line responses wrap correctly in cell
- [ ] Can process multiple cells in sequence

## Security Review

- [ ] No API keys in source code ✓
- [ ] No API keys in git history
- [ ] .gitignore covers .env, node_modules, etc.
- [ ] localStorage warning in README ✓
- [ ] HTTPS enforced in manifest (if production)
- [ ] No sensitive logs in console
- [ ] User input is escaped before display

## Performance & Optimization

- [ ] Page loads in < 2 seconds
- [ ] API responses return in 2-5 seconds
- [ ] Status indicators provide feedback during wait
- [ ] No memory leaks (check DevTools)
- [ ] CSS is minified (optional)
- [ ] JavaScript is not transpiled (using ES6+)

## Browser Compatibility

- [ ] Works in Edge (Chromium)
- [ ] Works in Chrome
- [ ] Works in Firefox
- [ ] Works in Safari
- [ ] localStorage works on all browsers

## GitHub Deployment (if using GitHub Pages)

### Repository Setup
- [ ] Created GitHub repository
- [ ] Initialized git: `git init`
- [ ] Added files: `git add .`
- [ ] Committed: `git commit -m "Initial commit"`
- [ ] Added remote: `git remote add origin https://github.com/username/codeexcel-ai.git`
- [ ] Pushed: `git push -u origin main`

### GitHub Pages Configuration
- [ ] Repository is public
- [ ] Went to Settings > Pages
- [ ] Selected "Deploy from a branch"
- [ ] Selected "main" branch and "/" folder
- [ ] Deployment is active
- [ ] Checked GitHub Pages URL is accessible

### Manifest Update for GitHub Pages
- [ ] Updated manifest.xml with GitHub Pages URL:
  ```
  https://username.github.io/codeexcel-ai/src/taskpane/html/taskpane.html
  ```
- [ ] Committed and pushed manifest changes
- [ ] Waited 30 seconds for GitHub Pages to rebuild

### Final URL Verification
- [ ] taskpane.html loads from GitHub Pages URL in browser
- [ ] Office.js library loads
- [ ] All JavaScript modules load
- [ ] CSS styling loads
- [ ] No 404 errors for assets

## Documentation Review

- [ ] README.md is complete and accurate
- [ ] OPENROUTER_API.md explains API integration
- [ ] Setup instructions are clear
- [ ] Troubleshooting section covers common issues
- [ ] API key storage process is documented
- [ ] All commands reference correct paths

## Final Testing Workflow

### Complete User Flow
1. [ ] Start fresh browser session (private mode)
2. [ ] Go to Excel Online
3. [ ] Insert the add-in using manifest URL
4. [ ] Add-in loads successfully
5. [ ] Go to Settings tab
6. [ ] Paste OpenRouter API key
7. [ ] Save settings
8. [ ] Refresh page - settings persist
9. [ ] Go to Assistant tab
10. [ ] Enter a test value in an Excel cell
11. [ ] Select the cell
12. [ ] Click Process Selection
13. [ ] Verify response appears in neighboring cell
14. [ ] Verify response box shows output
15. [ ] Check status indicators throughout

### Edge Cases
- [ ] Very long input (100+ characters)
- [ ] Special characters in input: @#$%^&*()
- [ ] Multiple rapid requests
- [ ] Internet disconnection mid-request
- [ ] API key with extra whitespace (should trim)
- [ ] Empty system prompt (use default)

## Production Readiness

- [ ] All critical errors have been resolved
- [ ] User-facing error messages are clear
- [ ] Status indicators work correctly
- [ ] API key is secure and persistent
- [ ] Documentation is complete
- [ ] No console warnings or errors
- [ ] Add-in functions as intended
- [ ] Ready for real users

## Monitoring & Maintenance

### Post-Deployment
- [ ] Monitor OpenRouter usage dashboard
- [ ] Check for error reports from users
- [ ] Monitor GitHub issues (if applicable)
- [ ] Track API response times
- [ ] Review Excel add-in telemetry (if available)

### Scheduled Tasks
- [ ] Weekly: Check OpenRouter account usage
- [ ] Monthly: Review API logs for errors
- [ ] Quarterly: Test all functionality
- [ ] Annually: Update dependencies and docs

## Known Issues & Workarounds

### Issue: localStorage blocked
**Workaround**: Use incognito/private mode or allow cookies for domain

### Issue: Add-in won't load
**Workaround**: Check manifest URL is accessible, verify HTTPS in prod

### Issue: Slow responses
**Workaround**: OpenRouter or Gemini may be overloaded - retry later

### Issue: API errors
**Workaround**: Verify API key is valid, check rate limits

---

## Sign-Off

- **Deployment Date**: _______________
- **Deployed By**: _______________
- **Tested By**: _______________
- **Environment**: [ ] Local [ ] Staging [ ] Production
- **Notes**: ___________________________________________________

### Deployment Status

- [ ] **READY FOR PRODUCTION** - All checks passed
- [ ] **NEEDS FIXES** - Issues identified (see notes)
- [ ] **BLOCKED** - Critical issues prevent deployment

---

**Checklist Version**: 1.0
**Last Updated**: January 2026
**Next Review**: [3 months]
