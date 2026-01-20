# CodeExcel AI - Architecture & Component Reference

Visual guide to how the Excel Add-in works.

## System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        EXCEL ONLINE                              │
│                    (office.com)                                  │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           │ Loads Task Pane
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│               OFFICE JAVASCRIPT API (CDN)                        │
│        https://appsforoffice.microsoft.com/lib/1/...            │
│  (Provides: Excel.run(), context.workbook, etc)                 │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                    TASK PANE HTML                                │
│               src/taskpane/html/taskpane.html                   │
└──────────────────┬──────────────────────┬───────────────────────┘
                   │                      │
        ┌──────────▼──────┐    ┌──────────▼──────┐
        │  Assistant Tab  │    │  Settings Tab   │
        ├─────────────────┤    ├─────────────────┤
        │ Process Button  │    │ API Key Input   │
        │ Response Box    │    │ Model Input     │
        │ Cell Display    │    │ Prompt Input    │
        │ Status Indicator│    │ Save Button     │
        └────────┬────────┘    └────────┬────────┘
                 │                      │
                 └──────────┬───────────┘
                            │
        ┌───────────────────▼───────────────────┐
        │      STYLESHEET (taskpane.css)        │
        │   Professional design, colors, layout │
        └───────────────────────────────────────┘

        ┌───────────────────┬───────────────────┐
        │                   │                   │
        ▼                   ▼                   ▼
    ┌─────────┐         ┌─────────┐       ┌──────────┐
    │storage. │         │   ui.js │       │  api.js  │
    │  js     │         │         │       │          │
    ├─────────┤         ├─────────┤       ├──────────┤
    │localStorage│       │Status   │       │ OpenRouter
    │wrapper  │         │Messages │       │ Integration
    │         │         │         │       │          │
    └────┬────┘         └────┬────┘       └────┬─────┘
         │                   │                 │
         └───────────────────┼─────────────────┘
                             │
                    ┌────────▼────────┐
                    │  taskpane.js    │
                    │ (Main Logic)    │
                    │                 │
                    │ - Tab switching │
                    │ - Cell selection│
                    │ - Button handlers
                    │ - Coordination  │
                    └─────────────────┘
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
        ▼                    ▼                    ▼
   ┌─────────┐          ┌─────────┐         ┌──────────┐
   │localStorage│        │ Excel   │         │OpenRouter│
   │(API Key,  │        │ Context │         │API       │
   │ Settings) │        │(Cells)  │         │(AI)      │
   └─────────┘          └─────────┘         └──────────┘
```

---

## Data Flow

### Processing a Cell (Happy Path)

```
User selects cell A1
        │
        ▼
taskpane.js: onSelectionChanged()
        │
        ├─ Read cell address → Display "A1"
        │
        ▼
User clicks "Process Selection"
        │
        ▼
taskpane.js: handleProcessSelection()
        │
        ├─ Check API key exists → Show error if not
        │
        ├─ UI: setProcessing("Reading cell...")
        │
        ├─ Excel.run(): Get selected range
        │  └─ cellValue = A1.values[0][0]
        │
        ├─ UI: setProcessing("Sending to AI...")
        │
        ├─ storage.js: Get stored settings
        │  └─ apiKey, model, systemPrompt
        │
        ├─ api.js: callOpenRouter()
        │  │
        │  ├─ Build request with system prompt
        │  │
        │  ├─ Fetch POST to OpenRouter
        │  │
        │  ├─ Parse response
        │  │
        │  └─ Return aiResponse
        │
        ├─ UI: displayResponse(aiResponse)
        │
        ├─ UI: setProcessing("Writing to Excel...")
        │
        ├─ Excel.run(): Write to target cell
        │  │
        │  ├─ targetRange = selection.getOffsetRange(0, 1)
        │  │
        │  ├─ targetRange.values = [[aiResponse]]
        │  │
        │  └─ targetRange.format.fill.color = "#E8F5E9"
        │
        ├─ UI: setSuccess("Response written to cell")
        │
        └─ Status auto-resets to Idle after 3 seconds
```

---

## Module Responsibilities

### storage.js
**Purpose**: Secure credential and settings management

```
Input → localStorage
│
├─ getApiKey() → Returns stored API key
├─ setApiKey(key) → Saves API key
├─ clearApiKey() → Removes API key
├─ hasApiKey() → Checks if key exists
│
├─ getModel() → Returns model (or default)
├─ setModel(model) → Saves model selection
│
├─ getSystemPrompt() → Returns prompt (or default)
├─ setSystemPrompt(prompt) → Saves prompt
│
└─ clearAll() → Removes everything
```

### api.js
**Purpose**: OpenRouter API communication

```
Input: (userMessage, apiKey, model, systemPrompt)
│
├─ Validate inputs (non-empty, valid key)
│
├─ Build request headers
│  └─ Authorization: Bearer {key}
│  └─ HTTP-Referer: https://codeexcel.ai
│  └─ X-Title: CodeExcel AI Assistant
│
├─ Build request body
│  ├─ model, messages, temperature, max_tokens
│  │
│  ├─ messages[0]: role="system", content=prompt
│  └─ messages[1]: role="user", content=message
│
├─ Fetch POST to OpenRouter
│
├─ Parse JSON response
│
├─ Extract: response.choices[0].message.content
│
└─ Output: AI response text
    └─ Or throw Error for failures
```

### ui.js
**Purpose**: User interface state management

```
Input: User actions & system state
│
├─ Status Management
│  ├─ setStatus(type, message)
│  ├─ setIdle() → "Ready"
│  ├─ setProcessing() → "Processing..."
│  ├─ setSuccess() → "Success" (3 sec auto-reset)
│  └─ setError() → Shows error message
│
├─ Response Display
│  ├─ displayResponse(text) → Shows in response box
│  ├─ clearResponse() → Resets to placeholder
│  └─ escapeHtml(text) → Safety measure
│
├─ Form Management
│  ├─ loadSettingsForm(settings) → Fill inputs
│  ├─ getSettingsForm() → Read form values
│  └─ showSettingsStatus(type, msg) → Feedback
│
└─ Utilities
   ├─ setProcessButtonEnabled(bool)
   └─ updateSelectedCell(address)
```

### taskpane.js
**Purpose**: Main application logic and coordination

```
Startup:
├─ initializeApp()
│  ├─ UIModule.init() → Cache DOM elements
│  │
│  ├─ Office.onReady() → Wait for Office.js
│  │
│  ├─ Load stored settings → Fill form
│  │
│  ├─ setupEventListeners() → Bind handlers
│  │
│  └─ trackExcelSelection() → Monitor changes
│
Event: Tab Button Click
├─ Update active tab button styling
└─ Show corresponding tab content
│
Event: Process Button Click
├─ handleProcessSelection()
│  ├─ Check API key exists
│  ├─ Read selected cell value
│  ├─ Call APIModule.callOpenRouter()
│  ├─ Display response
│  ├─ Write to adjacent cell
│  └─ Update status

Event: Save Settings Button Click
├─ handleSaveSettings()
│  ├─ Get form values
│  ├─ Validate API key
│  ├─ Save to storage
│  └─ Show confirmation

Event: Clear API Key Button Click
├─ handleClearApiKey()
│  ├─ Confirm with user
│  ├─ Clear storage
│  └─ Reset form
```

---

## State Management

### Browser Storage (localStorage)

```
localStorage = {
    "openrouter_api_key": "sk-...",
    "openrouter_model": "google/gemini-...",
    "openrouter_system_prompt": "You are an..."
}
```

- **Persistence**: Across browser sessions
- **Scope**: Single origin (domain + protocol + port)
- **Size Limit**: ~5-10MB per domain
- **Security**: Local browser only (not sent to server)

### In-Memory State

```
JavaScript Variables:
├─ currentSelectedRange → Current Excel selection
├─ UI.elements → Cached DOM references
├─ Form values → Temporary during editing
└─ API responses → Displayed then discarded
```

---

## Error Handling Flow

```
Operation starts
│
├─ Network Error?
│  ├─ "Unable to reach API"
│  └─ Check internet, OpenRouter status
│
├─ Missing API Key?
│  ├─ "API Key required - visit Settings tab"
│  └─ Cannot proceed
│
├─ Empty Cell?
│  ├─ "Selected cell is empty"
│  └─ Try different cell
│
├─ Invalid API Key?
│  ├─ "API Error: HTTP 401 Unauthorized"
│  └─ Update key in Settings
│
├─ Rate Limit?
│  ├─ "API Error: HTTP 429 Too Many Requests"
│  └─ Wait 1 minute before retry
│
├─ Invalid Response?
│  ├─ "Invalid API response format"
│  └─ Check model name, try refresh
│
└─ Success! ✅
   └─ Response written to cell
```

---

## Deployment Topology

### Local Development
```
Your Computer
    │
    ├─ npm start
    │  └─ http-server on port 3000
    │     └─ http://localhost:3000/taskpane.html
    │
    ├─ Browser
    │  └─ Loads taskpane.html
    │
    └─ ngrok (optional)
       └─ https://...ngrok.io/taskpane.html
          └─ Public tunnel for Excel Online
```

### GitHub Pages Deployment
```
GitHub Repository
    │
    ├─ /src/taskpane/html/taskpane.html
    │
    └─ GitHub Pages (automatic)
       └─ https://username.github.io/codeexcel-ai/src/taskpane/html/taskpane.html
          └─ manifest.xml points here
             └─ Excel Online loads add-in
```

---

## Sequence Diagram: Full Operation

```
User          UI          Excel.js       API.js        OpenRouter
 │             │             │             │               │
 ├─Select Cell──→            │             │               │
 │             │──load──→     │             │               │
 │             │←address──     │             │               │
 │             ├─Display──→   │             │               │
 │             │             │             │               │
 ├─Click Process─→            │             │               │
 │             │              │             │               │
 │             ├─Read Value──→ │             │               │
 │             │             │─getValues──→│               │
 │             │←values─────←─│             │               │
 │             ├─Show Processing            │               │
 │             │              │             │               │
 │             │              │             ├─Get Key────→  │
 │             │←─Key────────←─────────────│               │
 │             │              │             │               │
 │             │              │             ├─Build Req──→  │
 │             │              │             │               │
 │             │              │             ├─POST───────→  │
 │             │              │             │               │
 │             │              │             │             [Processing]
 │             │              │             │               │
 │             │              │             │←─Response───┤
 │             │              │             │               │
 │             ├─Show Response◄────────────┤               │
 │             │              │             │               │
 │             ├─Show Writing→             │               │
 │             │              ├─Write Cell│               │
 │             │              │─Success──→│               │
 │             ├─Show Success ←            │               │
 │             │              │             │               │
 └─Process Complete           │             │               │
```

---

## Component Interaction Map

```
┌──────────────────────────────────────────┐
│         taskpane.html                    │
│  (UI Container & Script Loader)          │
│                                          │
│  ┌────────────────────────────────────┐ │
│  │        taskpane.css                │ │
│  │    (Visual Styling)                │ │
│  └────────────────────────────────────┘ │
└──────────────────────────────────────────┘
             │
    ┌────────┼────────┬───────────┐
    │        │        │           │
    ▼        ▼        ▼           ▼
┌────────┐ ┌────────┐ ┌────────┐ ┌──────────┐
│storage │ │  ui.js │ │ api.js │ │taskpane.js
│  .js   │ │        │ │        │ │ (Main)   │
└────────┘ └────────┘ └────────┘ └──────────┘
    │        │        │           │
    └────────┼────────┼───────────┘
             │
    ┌────────┴────────┬────────────┐
    │                 │            │
    ▼                 ▼            ▼
┌─────────┐    ┌──────────┐   ┌────────┐
│Office.js│    │localStorage│   │Excel   │
│(API)    │    │(Storage)   │   │Online  │
└─────────┘    └──────────┘   └────────┘
    │                            │
    └────────┬───────────────────┘
             │
             ▼
         ┌────────────┐
         │OpenRouter  │
         │API (Gemini)│
         └────────────┘
```

---

## User Journey Map

```
1. User Opens Excel Online
   ↓
2. Goes to Insert → Office Add-ins → Upload My Add-in
   ↓
3. Selects manifest.xml
   ↓
4. Task Pane Opens (Loads taskpane.html)
   ↓
5. Sees "Settings" tab recommendation
   ↓
6. Clicks Settings Tab
   ↓
7. Pastes OpenRouter API Key
   ↓
8. Clicks "Save Settings"
   ↓
9. Returns to "Assistant" Tab
   ↓
10. Selects Excel cell with data
    ↓
11. Clicks "Process Selection"
    ↓
12. Sees "Processing..." status
    ↓
13. AI response appears in response box
    ↓
14. Response also written to adjacent cell
    ↓
15. Sees "Success" status
    ↓
16. Can process more cells or adjust settings
```

---

## Security Model

```
┌──────────────────────┐
│   User's Browser     │
│                      │
│  ┌────────────────┐  │
│  │   localStorage │  │
│  │                │  │
│  │ API Key (User) │  │──┐ User-specific
│  │ Settings       │  │  │ Local storage
│  └────────────────┘  │  │ Never sent to repo
│                      │  │
│  ┌────────────────┐  │  │
│  │  Session RAM   │  │  │ Temporary memory
│  │  API Response  │  │  │ Cleared on refresh
│  └────────────────┘  │  │
│                      │  │
└──────────────────────┘  │
         │                │
         ├─HTTPS only─→   │
         │                │
         ▼                │
    ┌─────────┐           │
    │OpenRouter│◄──────────┘
    │ (Remote)│
    │         │
    │ No Keys │
    │ Stored  │
    └─────────┘
```

---

**Document**: Architecture Reference  
**Version**: 1.0  
**Updated**: January 2026  
**Status**: Production
