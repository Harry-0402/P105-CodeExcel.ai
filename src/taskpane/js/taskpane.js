/**
 * Task Pane Module: Main application logic with Gemini-style conversational UI
 * Coordinates tab navigation, message handling, and Excel integration
 */

const TaskPane = {
    // State
    currentSelectedCell: null,
    messageHistory: [],
    isProcessing: false,
    autoSwitchEnabled: true,
    cacheEnabled: true,

    /**
     * Initialize the application
     */
    async init() {
        console.log('Initializing TaskPane...');
        
        // Check if Office.js is available
        if (typeof Office !== 'undefined') {
            // Wait for Office.js to be ready
            await Office.onReady(async (info) => {
                if (info.host !== Office.HostType.Excel) {
                    this.showError('This add-in only works with Excel');
                    return;
                }

                console.log('Excel Office.js is ready');

                // Initialize sub-modules
                StorageModule.init();
                ModelsManager.init();
                
                // Load settings
                this.loadSettings();

                // Setup event listeners
                this.setupEventListeners();

                // Setup Excel event tracking
                this.setupExcelTracking();

                // Initialize with welcome state
                this.showWelcome();

                console.log('TaskPane initialization complete');
            });
        } else {
            // Running in browser without Office.js (for testing)
            console.log('Running in browser mode (Office.js not available)');
            
            // Initialize sub-modules if available
            if (typeof StorageModule !== 'undefined') StorageModule.init();
            if (typeof ModelsManager !== 'undefined') ModelsManager.init();
            
            // Load settings
            this.loadSettings();

            // Setup event listeners
            this.setupEventListeners();

            // Initialize with welcome state
            this.showWelcome();

            console.log('Browser mode initialization complete');
        }
    },

    /**
     * Load saved settings
     */
    loadSettings() {
        const settings = StorageModule.getAllSettings();
        
        // Load API key
        const apiKeyInput = document.getElementById('apiKey');
        if (apiKeyInput && settings.apiKey) {
            apiKeyInput.value = settings.apiKey;
        }

        // Load model selection
        const modelSelect = document.getElementById('modelSelect');
        if (modelSelect && settings.model) {
            modelSelect.value = settings.model;
            if (typeof ModelsManager !== 'undefined') {
                ModelsManager.setModel(settings.model);
            }
        }

        // Load temperature
        const tempSlider = document.getElementById('tempSlider');
        const tempValue = document.getElementById('tempValue');
        if (tempSlider && settings.temperature) {
            tempSlider.value = settings.temperature;
            if (tempValue) {
                tempValue.textContent = settings.temperature;
            }
        }

        // Load preferences
        this.autoSwitchEnabled = settings.autoSwitch !== false;
        this.cacheEnabled = settings.cacheResponses !== false;

        const autoSwitchCheckbox = document.getElementById('autoSwitchCheckbox');
        const cacheResponsesCheckbox = document.getElementById('cacheResponsesCheckbox');
        if (autoSwitchCheckbox) autoSwitchCheckbox.checked = this.autoSwitchEnabled;
        if (cacheResponsesCheckbox) cacheResponsesCheckbox.checked = this.cacheEnabled;
    },

    /**
     * Setup all event listeners
     */
    setupEventListeners() {
        console.log('Setting up event listeners...');
        
        // Tab navigation
        const tabButtons = document.querySelectorAll('.tab-button');
        console.log('Found tab buttons:', tabButtons.length);
        tabButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                const tab = e.target.closest('.tab-button').dataset.tab;
                console.log('Tab clicked:', tab);
                this.switchTab(tab);
            });
        });

        // Message input
        const messageInput = document.getElementById('messageInput');
        if (messageInput) {
            messageInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    this.sendMessage();
                }
            });
        }

        // Send button
        document.getElementById('sendButton').addEventListener('click', () => this.sendMessage());

        // Suggestion cards
        document.querySelectorAll('.suggestion-card').forEach(card => {
            card.addEventListener('click', () => this.handleSuggestion(card.dataset.action));
        });

        // Dark mode toggle
        document.getElementById('darkModeToggle').addEventListener('change', () => this.toggleDarkMode());

        // Settings
        document.getElementById('saveSettings').addEventListener('click', () => this.saveSettings());
        document.getElementById('resetSettings').addEventListener('click', () => this.resetSettings());
        document.getElementById('tempSlider').addEventListener('input', (e) => {
            document.getElementById('tempValue').textContent = e.target.value;
        });

        // History
        document.getElementById('clearHistoryBtn').addEventListener('click', () => this.clearHistory());
        document.getElementById('historySearch').addEventListener('input', (e) => this.searchHistory(e.target.value));

        // Password visibility toggle
        document.getElementById('toggleApiKeyVisibility').addEventListener('click', () => {
            const input = document.getElementById('apiKey');
            input.type = input.type === 'password' ? 'text' : 'password';
        });

        // Model selection
        document.getElementById('modelSelect').addEventListener('change', (e) => {
            const model = e.target.value;
            if (model === 'auto') {
                // Keep auto-switch behavior
            } else {
                ModelsManager.setModel(model);
            }
        });

        console.log('Event listeners setup complete');
    },

    /**
     * Setup Excel event tracking
     */
    setupExcelTracking() {
        try {
            Excel.run(async (context) => {
                // Track cell/range selection changes
                context.application.onSelectionChanged.add(() => {
                    this.updateSelectedCell(context);
                });

                await context.sync();
                
                // Initial selection check
                this.updateSelectedCell(context);
            }).catch(err => {
                console.error('Excel tracking setup error:', err);
            });
        } catch (err) {
            console.error('Excel tracking setup error:', err);
        }
    },

    /**
     * Update selected cell display
     */
    async updateSelectedCell(context) {
        try {
            const range = context.application.getSelectedData(Excel.ValueFormat.values);
            const address = (await range.load('address')).address;
            
            this.currentSelectedCell = address;

            const tag = document.getElementById('selectedCellTag');
            const value = document.getElementById('selectedCellValue');
            if (tag && value) {
                value.textContent = address;
                tag.style.display = 'inline-flex';
            }

            await context.sync();
        } catch (err) {
            console.log('Cell selection tracking:', err.message);
        }
    },

    /**
     * Switch to a different tab
     */
    switchTab(tabName) {
        // Update tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // Update tab panes
        document.querySelectorAll('.tab-pane').forEach(pane => {
            pane.classList.remove('active');
        });
        document.getElementById(tabName).classList.add('active');

        // Load history when switching to history tab
        if (tabName === 'history') {
            this.loadHistory();
        }
    },

    /**
     * Send a message to the AI
     */
    async sendMessage() {
        const messageInput = document.getElementById('messageInput');
        const message = messageInput.value.trim();

        if (!message) return;

        // Get API key
        const apiKey = document.getElementById('apiKey').value;
        if (!apiKey) {
            this.showError('API Key required. Please add it in Settings.');
            return;
        }

        // Get model
        const modelSelect = document.getElementById('modelSelect');
        const model = modelSelect.value === 'auto' ? 'auto' : modelSelect.value;

        // Get temperature
        const temperature = parseFloat(document.getElementById('tempSlider').value);

        // Clear input
        messageInput.value = '';
        messageInput.disabled = true;

        // Update status
        this.setProcessing(true);

        try {
            // Add user message to UI
            this.addMessage('user', message);

            // Get selected cell content for context
            let cellContext = '';
            if (this.currentSelectedCell) {
                cellContext = `Selected cell: ${this.currentSelectedCell}\n`;
            }

            // Make API call
            const response = await APIModule.callOpenRouter(
                cellContext + message,
                apiKey,
                model,
                'You are a helpful assistant for Excel data processing and analysis.',
                { temperature }
            );

            // Add AI response
            this.addMessage('ai', response.response);

            // Save to history
            this.saveToHistory(message, response.response, response.model);

            // Update UI with model info
            if (response.model) {
                const modelInfo = ModelsManager.getModelInfo(response.model);
                if (modelInfo) {
                    const badge = document.getElementById('currentModel');
                    badge.textContent = modelInfo.name.split(' ')[0];
                }
            }

            this.setProcessing(false);
        } catch (error) {
            console.error('API error:', error);
            this.showError(`Error: ${error.message}`);
            this.setProcessing(false);
        }

        messageInput.disabled = false;
        messageInput.focus();
    },

    /**
     * Handle suggestion card clicks
     */
    handleSuggestion(action) {
        const suggestions = {
            'summarize': 'Summarize the selected data in a concise format.',
            'extract': 'Extract the key information from the selected data.',
            'analyze': 'Analyze the selected data and provide insights.',
            'format': 'Format the selected data into a structured format.'
        };

        const message = suggestions[action];
        if (message) {
            const messageInput = document.getElementById('messageInput');
            messageInput.value = message;
            this.sendMessage();
        }
    },

    /**
     * Add a message to the conversation
     */
    addMessage(sender, text) {
        const messageArea = document.getElementById('messageArea');
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${sender}`;
        messageDiv.innerHTML = `<div class="message-content">${this.escapeHtml(text)}</div>`;
        messageArea.appendChild(messageDiv);
        messageArea.scrollTop = messageArea.scrollHeight;

        this.messageHistory.push({ sender, text, timestamp: new Date() });

        // Hide welcome section after first message
        const welcomeSection = messageArea.querySelector('.welcome-section');
        if (welcomeSection) {
            welcomeSection.remove();
        }
    },

    /**
     * Show welcome section
     */
    showWelcome() {
        const messageArea = document.getElementById('messageArea');
        if (messageArea.children.length === 1) {
            // Already showing welcome
            return;
        }
        
        const welcomeDiv = document.createElement('div');
        welcomeDiv.className = 'welcome-section';
        welcomeDiv.innerHTML = `
            <h2>Welcome to CodeExcel AI</h2>
            <p>Select a cell or range in Excel and ask me to help you analyze, transform, or process data.</p>
            <div class="quick-tips">
                <div class="tip-item">
                    <i class="fas fa-lightbulb"></i>
                    <p><strong>Tip:</strong> Use @ to reference cells or ranges</p>
                </div>
            </div>
        `;
        messageArea.innerHTML = '';
        messageArea.appendChild(welcomeDiv);
    },

    /**
     * Save message to history
     */
    saveToHistory(question, answer, model) {
        const history = StorageModule.getHistory() || [];
        history.unshift({
            question,
            answer,
            model,
            timestamp: new Date().toISOString()
        });
        
        // Keep only last 50 items
        if (history.length > 50) {
            history.pop();
        }
        
        StorageModule.setHistory(history);
    },

    /**
     * Load history from storage
     */
    loadHistory() {
        const history = StorageModule.getHistory() || [];
        const historyList = document.getElementById('historyList');

        if (history.length === 0) {
            historyList.innerHTML = `
                <div class="history-empty">
                    <i class="fas fa-inbox"></i>
                    <p>No history yet</p>
                </div>
            `;
            return;
        }

        historyList.innerHTML = '';
        history.forEach(item => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'history-item';
            const date = new Date(item.timestamp);
            itemDiv.innerHTML = `
                <div class="history-item-text">${this.escapeHtml(item.question)}</div>
                <div class="history-item-time">${date.toLocaleString()} • ${item.model}</div>
            `;
            itemDiv.addEventListener('click', () => {
                document.getElementById('messageInput').value = item.question;
            });
            historyList.appendChild(itemDiv);
        });
    },

    /**
     * Search history
     */
    searchHistory(query) {
        const history = StorageModule.getHistory() || [];
        const historyList = document.getElementById('historyList');
        const filtered = history.filter(item => 
            item.question.toLowerCase().includes(query.toLowerCase()) ||
            item.answer.toLowerCase().includes(query.toLowerCase())
        );

        historyList.innerHTML = '';
        filtered.forEach(item => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'history-item';
            const date = new Date(item.timestamp);
            itemDiv.innerHTML = `
                <div class="history-item-text">${this.escapeHtml(item.question)}</div>
                <div class="history-item-time">${date.toLocaleString()}</div>
            `;
            historyList.appendChild(itemDiv);
        });
    },

    /**
     * Clear all history
     */
    clearHistory() {
        if (confirm('Are you sure you want to clear all history?')) {
            StorageModule.setHistory([]);
            this.loadHistory();
        }
    },

    /**
     * Save settings
     */
    saveSettings() {
        const apiKey = document.getElementById('apiKey').value;
        const model = document.getElementById('modelSelect').value;
        const temperature = parseFloat(document.getElementById('tempSlider').value);
        const autoSwitch = document.getElementById('autoSwitchCheckbox').checked;
        const cacheResponses = document.getElementById('cacheResponsesCheckbox').checked;

        StorageModule.setApiKey(apiKey);
        StorageModule.setModel(model);
        StorageModule.setSetting('temperature', temperature);
        StorageModule.setSetting('autoSwitch', autoSwitch);
        StorageModule.setSetting('cacheResponses', cacheResponses);

        this.autoSwitchEnabled = autoSwitch;
        this.cacheEnabled = cacheResponses;

        this.showSuccess('Settings saved!');
    },

    /**
     * Reset settings to default
     */
    resetSettings() {
        if (confirm('Reset settings to default?')) {
            StorageModule.clearAllSettings();
            this.loadSettings();
            this.showSuccess('Settings reset to default');
        }
    },

    /**
     * Toggle dark mode
     */
    toggleDarkMode() {
        document.body.classList.toggle('dark-mode');
        const isDarkMode = document.body.classList.contains('dark-mode');
        StorageModule.setSetting('darkMode', isDarkMode);
    },

    /**
     * Set processing state
     */
    setProcessing(isProcessing) {
        this.isProcessing = isProcessing;
        const sendBtn = document.getElementById('sendButton');
        const statusBadge = document.getElementById('statusIndicator');

        if (isProcessing) {
            sendBtn.disabled = true;
            statusBadge.textContent = 'Processing...';
            statusBadge.className = 'status-badge status-processing';
        } else {
            sendBtn.disabled = false;
            statusBadge.textContent = 'Ready';
            statusBadge.className = 'status-badge status-ready';
        }
    },

    /**
     * Show error message
     */
    showError(message) {
        const statusBadge = document.getElementById('statusIndicator');
        statusBadge.textContent = 'Error';
        statusBadge.className = 'status-badge status-error';
        console.error(message);
        alert(message);
    },

    /**
     * Show success message
     */
    showSuccess(message) {
        const statusBadge = document.getElementById('statusIndicator');
        statusBadge.textContent = 'Success';
        statusBadge.className = 'status-badge status-ready';
        console.log(message);
    },

    /**
     * Escape HTML to prevent XSS
     */
    escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, m => map[m]);
    }
};

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => TaskPane.init());
} else {
    TaskPane.init();
}
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            const tabName = e.target.getAttribute('data-tab');

            // Update active states
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            e.target.classList.add('active');
            document.getElementById(tabName)?.classList.add('active');
        });
    });

    // Process button
    const processButton = document.getElementById('processButton');
    if (processButton) {
        processButton.addEventListener('click', handleProcessSelection);
    }

    // Save settings button
    const saveButton = document.getElementById('saveButton');
    if (saveButton) {
        saveButton.addEventListener('click', handleSaveSettings);
    }

    // Clear API key button
    const clearButton = document.getElementById('clearButton');
    if (clearButton) {
        clearButton.addEventListener('click', handleClearApiKey);
    }
}

/**
 * Track Excel cell selection and update UI
 */
async function trackExcelSelection() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onSelectionChanged.add(onSelectionChanged);
            await context.sync();
        });
    } catch (error) {
        console.error('Error setting up selection tracking:', error);
    }

    // Initial update
    updateSelectedCellDisplay();
}

/**
 * Handle selection change event
 */
async function onSelectionChanged() {
    updateSelectedCellDisplay();
}

/**
 * Update the selected cell display
 */
async function updateSelectedCellDisplay() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedData(Excel.ValueType.values);
            const cellAddress = context.application.getSelection();
            
            // Get the currently selected range
            currentSelectedRange = context.workbook.getSelectedData(Excel.ValueType.values);
            
            // Update UI with cell address (simplified - shows last part)
            const activeCell = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            const selection = context.workbook.worksheets.getActiveWorksheet().getRange('A1').getOffsetRange(0, 0).getIntersection(context.application.getSelection());
            
            // Try to get the selection address
            try {
                const selectedRange = context.workbook.worksheets.getActiveWorksheet().getSelection();
                selectedRange.load('address');
                await context.sync();
                UIModule.updateSelectedCell(selectedRange.address);
            } catch (err) {
                // Fallback if address is not available
                console.log('Using fallback cell address');
                UIModule.updateSelectedCell('Selected');
            }
        });
    } catch (error) {
        console.error('Error updating selected cell:', error);
    }
}

/**
 * Handle Process Selection button click
 */
async function handleProcessSelection() {
    // Check if API key exists
    if (!StorageModule.hasApiKey()) {
        UIModule.setError('API Key required - visit Settings tab');
        return;
    }

    UIModule.setProcessing('Reading cell...');

    try {
        let cellValue = null;
        let targetRange = null;

        await Excel.run(async (context) => {
            // Get the selected range
            const range = context.workbook.worksheets.getActiveWorksheet().getSelection();
            range.load('values,address');
            await context.sync();

            cellValue = range.values[0][0];
            const cellAddress = range.address;

            // Get the offset range (one column to the right)
            targetRange = range.getOffsetRange(0, 1);

            if (!cellValue) {
                throw new Error('Selected cell is empty');
            }

            console.log(`Processing cell ${cellAddress} with value: ${cellValue}`);
        });

        // Validate cell value
        if (cellValue === null || cellValue === undefined || cellValue === '') {
            UIModule.setError('Selected cell is empty');
            return;
        }

        UIModule.setProcessing('Sending to AI...');

        // Get settings
        const apiKey = StorageModule.getApiKey();
        const model = StorageModule.getModel();
        const systemPrompt = StorageModule.getSystemPrompt();

        // Call API
        const response = await APIModule.callOpenRouter(
            cellValue.toString(),
            apiKey,
            model,
            systemPrompt
        );

        // Display response
        UIModule.displayResponse(response);
        UIModule.setProcessing('Writing to Excel...');

        // Write response to target cell
        await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getSelection();
            const targetRange = range.getOffsetRange(0, 1);
            
            targetRange.values = [[response]];
            targetRange.format.fill.color = '#E8F5E9';
            
            await context.sync();
        });

        UIModule.setSuccess('Response written to cell');
    } catch (error) {
        const errorMessage = APIModule.formatErrorMessage(error);
        UIModule.setError(errorMessage);
        console.error('Error processing selection:', error);
    }
}

/**
 * Handle Save Settings button click
 */
async function handleSaveSettings() {
    const settings = UIModule.getSettingsForm();

    // Validate input
    if (!settings.apiKey || settings.apiKey.trim() === '') {
        UIModule.showSettingsStatus('error', 'API Key cannot be empty');
        return;
    }

    // Save to storage
    const apiKeySaved = StorageModule.setApiKey(settings.apiKey);
    const modelSaved = StorageModule.setModel(settings.model);
    const promptSaved = StorageModule.setSystemPrompt(settings.systemPrompt);

    if (apiKeySaved && modelSaved && promptSaved) {
        UIModule.showSettingsStatus('success', 'Settings saved successfully');
        UIModule.setProcessButtonEnabled(true);
        UIModule.setIdle('API Key configured');
    } else {
        UIModule.showSettingsStatus('error', 'Failed to save settings');
    }
}

/**
 * Handle Clear API Key button click
 */
async function handleClearApiKey() {
    const confirmed = confirm('Are you sure you want to clear the API Key?');
    
    if (confirmed) {
        const cleared = StorageModule.clearApiKey();
        
        if (cleared) {
            UIModule.elements.apiKeyInput.value = '';
            UIModule.showSettingsStatus('success', 'API Key cleared');
            UIModule.setProcessButtonEnabled(false);
            UIModule.setError('API Key removed - add a new one to continue');
        } else {
            UIModule.showSettingsStatus('error', 'Failed to clear API Key');
        }
    }
}

/**
 * Initialize when Office.js is ready
 */
TaskPane.init();
