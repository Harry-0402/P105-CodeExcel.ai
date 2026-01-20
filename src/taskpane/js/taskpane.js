/**
 * CodeExcel.AI - Task Pane Main Module
 * Clean, simple initialization for both browser and Excel
 */

const TaskPane = {
    currentSelectedCell: null,
    messageHistory: [],
    isProcessing: false,
    inlineAssistEnabled: true,
    currentFormulaSuggestion: null,
    lastInlineTrigger: null,
    isApplyingFormula: false,

    /**
     * Initialize the application
     */
    init() {
        console.log('🚀 Initializing CodeExcel.AI...');
        
        // Setup basic UI first
        this.setupEventListeners();
        this.loadSettings();
        this.showWelcome();
        
        // Initialize Office.js features if available
        if (typeof Office !== 'undefined') {
            Office.onReady((info) => {
                if (info.host === Office.HostType.Excel) {
                    console.log('✅ Excel Office.js ready');
                    this.setupExcelTracking();
                }
            });
        } else {
            console.log('ℹ️ Running in browser mode (Office.js not available)');
        }
        
        // Initialize modules if available
        if (typeof StorageModule !== 'undefined') StorageModule.init();
        if (typeof ModelsManager !== 'undefined') ModelsManager.init();
        
        console.log('✅ TaskPane initialized');
    },

    /**
     * Setup all event listeners
     */
    setupEventListeners() {
        console.log('🔧 Setting up event listeners...');
        
        // Tab navigation
        const tabButtons = document.querySelectorAll('.tab-button');
        console.log(`   Found ${tabButtons.length} tab buttons`);
        
        tabButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                const target = e.currentTarget;
                const tab = target.dataset.tab;
                console.log(`   Tab clicked: ${tab}`);
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
        const sendButton = document.getElementById('sendButton');
        if (sendButton) {
            sendButton.addEventListener('click', () => this.sendMessage());
        }

        // Suggestion cards
        document.querySelectorAll('.suggestion-card').forEach(card => {
            card.addEventListener('click', () => this.handleSuggestion(card.dataset.action));
        });

        // Settings buttons
        const saveSettings = document.getElementById('saveSettings');
        if (saveSettings) {
            saveSettings.addEventListener('click', () => this.saveSettings());
        }

        const resetSettings = document.getElementById('resetSettings');
        if (resetSettings) {
            resetSettings.addEventListener('click', () => this.resetSettings());
        }

        // Inline assist actions
        const assistAccept = document.getElementById('formulaAssistAccept');
        if (assistAccept) {
            assistAccept.addEventListener('click', () => this.applyFormulaSuggestion());
        }

        const assistDismiss = document.getElementById('formulaAssistDismiss');
        if (assistDismiss) {
            assistDismiss.addEventListener('click', () => this.hideFormulaAssist());
        }

        // Temperature slider
        const tempSlider = document.getElementById('tempSlider');
        if (tempSlider) {
            tempSlider.addEventListener('input', (e) => {
                const tempValue = document.getElementById('tempValue');
                if (tempValue) {
                    tempValue.textContent = e.target.value;
                }
            });
        }

        // Dark mode toggle
        const darkModeToggle = document.getElementById('darkModeToggle');
        if (darkModeToggle) {
            darkModeToggle.addEventListener('change', () => this.toggleDarkMode());
        }

        // History buttons
        const clearHistoryBtn = document.getElementById('clearHistoryBtn');
        if (clearHistoryBtn) {
            clearHistoryBtn.addEventListener('click', () => this.clearHistory());
        }

        const historySearch = document.getElementById('historySearch');
        if (historySearch) {
            historySearch.addEventListener('input', (e) => this.searchHistory(e.target.value));
        }

        // API key visibility toggle
        const toggleApiKeyVisibility = document.getElementById('toggleApiKeyVisibility');
        if (toggleApiKeyVisibility) {
            toggleApiKeyVisibility.addEventListener('click', () => {
                const input = document.getElementById('apiKey');
                if (input) {
                    input.type = input.type === 'password' ? 'text' : 'password';
                }
            });
        }

        // Model selection
        const modelSelect = document.getElementById('modelSelect');
        if (modelSelect) {
            modelSelect.addEventListener('change', (e) => {
                const model = e.target.value;
                if (model !== 'auto' && typeof ModelsManager !== 'undefined') {
                    ModelsManager.setModel(model);
                }
            });
        }

        const inlineAssistToggle = document.getElementById('inlineAssistToggle');
        if (inlineAssistToggle) {
            inlineAssistToggle.addEventListener('change', (e) => {
                this.inlineAssistEnabled = e.target.checked;
                if (typeof StorageModule !== 'undefined') {
                    StorageModule.setSetting('inlineAssist', this.inlineAssistEnabled);
                }
                if (!this.inlineAssistEnabled) {
                    this.hideFormulaAssist();
                }
            });
        }

        console.log('✅ Event listeners attached');
    },

    /**
     * Switch between tabs
     */
    switchTab(tabName) {
        console.log(`🔄 Switching to tab: ${tabName}`);
        
        // Update tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => {
            btn.classList.remove('active');
        });
        const activeButton = document.querySelector(`[data-tab="${tabName}"]`);
        if (activeButton) {
            activeButton.classList.add('active');
        }

        // Update tab panes
        document.querySelectorAll('.tab-pane').forEach(pane => {
            pane.classList.remove('active');
        });
        const activePane = document.getElementById(tabName);
        if (activePane) {
            activePane.classList.add('active');
        }

        // Load history when switching to history tab
        if (tabName === 'history') {
            this.loadHistory();
        }
        
        console.log('✅ Tab switched');
    },

    /**
     * Load saved settings
     */
    loadSettings() {
        if (typeof StorageModule === 'undefined') return;
        
        const settings = StorageModule.getAllSettings();
        
        const apiKeyInput = document.getElementById('apiKey');
        if (apiKeyInput && settings.apiKey) {
            apiKeyInput.value = settings.apiKey;
        }

        const modelSelect = document.getElementById('modelSelect');
        if (modelSelect && settings.model) {
            modelSelect.value = settings.model;
        }

        const tempSlider = document.getElementById('tempSlider');
        const tempValue = document.getElementById('tempValue');
        if (tempSlider && settings.temperature) {
            tempSlider.value = settings.temperature;
            if (tempValue) {
                tempValue.textContent = settings.temperature;
            }
        }

        const autoSwitchCheckbox = document.getElementById('autoSwitchCheckbox');
        if (autoSwitchCheckbox) {
            autoSwitchCheckbox.checked = settings.autoSwitch !== false;
        }

        const cacheResponsesCheckbox = document.getElementById('cacheResponsesCheckbox');
        if (cacheResponsesCheckbox) {
            cacheResponsesCheckbox.checked = settings.cacheResponses !== false;
        }

        const inlineAssistToggle = document.getElementById('inlineAssistToggle');
        this.inlineAssistEnabled = settings.inlineAssist !== false;
        if (inlineAssistToggle) {
            inlineAssistToggle.checked = this.inlineAssistEnabled;
        }
    },

    /**
     * Show welcome message
     */
    showWelcome() {
        console.log('👋 Showing welcome');
        // Welcome section is already in HTML
    },

    /**
     * Setup Excel event tracking
     */
    setupExcelTracking() {
        try {
            Excel.run(async (context) => {
                context.workbook.onSelectionChanged.add(() => {
                    this.updateSelectedCell();
                });

                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.onChanged.add((eventArgs) => {
                    this.handleWorksheetChange(eventArgs);
                });
                await context.sync();
                this.updateSelectedCell();
            });
        } catch (err) {
            console.error('Excel tracking setup error:', err);
        }
    },

    /**
     * Update selected cell display
     */
    async updateSelectedCell() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load('address');
                await context.sync();
                
                this.currentSelectedCell = range.address;
                
                const tag = document.getElementById('selectedCellTag');
                const value = document.getElementById('selectedCellValue');
                if (tag && value) {
                    value.textContent = range.address;
                    tag.classList.remove('hidden');
                }
            });
        } catch (err) {
            console.log('Cell selection:', err.message);
        }
    },

    /**
     * Handle worksheet edits to detect formulas and trigger inline assist
     */
    async handleWorksheetChange(eventArgs) {
        if (!this.inlineAssistEnabled || typeof Office === 'undefined') return;
        if (this.isApplyingFormula) return;

        try {
            await Excel.run(async (context) => {
                const sheet = eventArgs?.worksheetId
                    ? context.workbook.worksheets.getItem(eventArgs.worksheetId)
                    : context.workbook.worksheets.getActiveWorksheet();

                const address = eventArgs?.address || this.currentSelectedCell || 'A1';
                const range = sheet.getRange(address);
                range.load(['address', 'formulas']);
                await context.sync();

                const formula = range.formulas?.[0]?.[0];
                if (!formula || !formula.toString().trim().startsWith('=')) {
                    return;
                }

                if (this.lastInlineTrigger && this.lastInlineTrigger.address === range.address) {
                    const elapsed = Date.now() - this.lastInlineTrigger.time;
                    if (elapsed < 1500) {
                        return;
                    }
                }

                this.lastInlineTrigger = { address: range.address, time: Date.now() };
                await this.requestFormulaSuggestion({ address: range.address, userFormula: formula });
            });
        } catch (err) {
            console.error('Worksheet change handler error:', err);
        }
    },

    /**
     * Extract comprehensive data from Excel
     */
    async extractExcelData() {
        if (typeof Office === 'undefined') {
            return null;
        }

        try {
            return await Excel.run(async (context) => {
                const data = {
                    selectedRange: null,
                    selectedValues: null,
                    tables: [],
                    usedRange: null,
                    sheetName: null
                };

                // Get active worksheet
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load('name');

                // Get selected range
                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);

                // Get used range (all data in sheet)
                const usedRange = sheet.getUsedRange();
                usedRange.load(['address', 'values', 'rowCount', 'columnCount']);

                // Get all tables in the worksheet
                const tables = sheet.tables;
                tables.load(['name', 'items']);

                await context.sync();

                // Extract sheet name
                data.sheetName = sheet.name;

                // Extract selected range data
                data.selectedRange = {
                    address: selectedRange.address,
                    rowCount: selectedRange.rowCount,
                    columnCount: selectedRange.columnCount,
                    values: selectedRange.values,
                    formulas: selectedRange.formulas
                };

                // Extract used range (limit to reasonable size)
                if (usedRange.rowCount <= 100 && usedRange.columnCount <= 20) {
                    data.usedRange = {
                        address: usedRange.address,
                        rowCount: usedRange.rowCount,
                        columnCount: usedRange.columnCount,
                        values: usedRange.values
                    };
                } else {
                    // For large datasets, only include headers and first few rows
                    const limitedRange = sheet.getRangeByIndexes(0, 0, Math.min(usedRange.rowCount, 10), usedRange.columnCount);
                    limitedRange.load(['address', 'values']);
                    await context.sync();
                    
                    data.usedRange = {
                        address: usedRange.address,
                        rowCount: usedRange.rowCount,
                        columnCount: usedRange.columnCount,
                        values: limitedRange.values,
                        truncated: true,
                        note: `Dataset has ${usedRange.rowCount} rows. Showing first 10 rows.`
                    };
                }

                // Extract table data
                for (let i = 0; i < tables.items.length; i++) {
                    const table = tables.items[i];
                    table.load(['name']);
                    
                    const tableRange = table.getRange();
                    tableRange.load(['address', 'values']);
                    
                    const headerRange = table.getHeaderRowRange();
                    headerRange.load('values');
                    
                    await context.sync();
                    
                    data.tables.push({
                        name: table.name,
                        address: tableRange.address,
                        headers: headerRange.values[0],
                        values: tableRange.values,
                        rowCount: tableRange.values.length - 1 // Exclude header
                    });
                }

                console.log('📊 Extracted Excel data:', data);
                return data;
            });
        } catch (error) {
            console.error('Error extracting Excel data:', error);
            return null;
        }
    },

    /**
     * Request a formula suggestion after detecting a user-entered '=' formula
     */
    async requestFormulaSuggestion({ address, userFormula }) {
        if (!this.inlineAssistEnabled) return;
        if (!address || !userFormula) return;
        if (typeof APIModule === 'undefined') return;

        const apiKeyInput = document.getElementById('apiKey');
        const apiKey = apiKeyInput ? apiKeyInput.value : '';
        if (!apiKey) return;

        const modelSelect = document.getElementById('modelSelect');
        const model = modelSelect ? modelSelect.value : 'auto';
        const tempSlider = document.getElementById('tempSlider');
        const temperature = tempSlider ? parseFloat(tempSlider.value) : 0.7;

        try {
            const excelData = await this.extractExcelData();
            const suggestion = await APIModule.getFormulaSuggestion({
                apiKey,
                model,
                temperature: Math.min(temperature, 1.0),
                address,
                userFormula,
                excelData
            });

            if (suggestion?.success && suggestion.formula) {
                this.currentFormulaSuggestion = { address, formula: suggestion.formula };
                this.showFormulaAssist(suggestion.formula);
            }
        } catch (error) {
            console.error('Formula suggestion error:', error);
        }
    },

    /**
     * Send a message to the AI
     */
    async sendMessage() {
        const messageInput = document.getElementById('messageInput');
        const message = messageInput.value.trim();

        if (!message) return;

        console.log('📤 Sending message:', message);

        // Get API key
        const apiKeyInput = document.getElementById('apiKey');
        const apiKey = apiKeyInput ? apiKeyInput.value : '';
        
        if (!apiKey) {
            this.addMessage('ai', '⚠️ Please add your OpenRouter API key in Settings tab');
            this.switchTab('settings');
            return;
        }

        // Clear input
        messageInput.value = '';
        this.setProcessing(true);

        // Add user message to chat
        this.addMessage('user', message);

        try {
            // Get model and temperature
            const modelSelect = document.getElementById('modelSelect');
            const model = modelSelect ? modelSelect.value : 'auto';
            
            const tempSlider = document.getElementById('tempSlider');
            const temperature = tempSlider ? parseFloat(tempSlider.value) : 0.7;

            // Extract Excel data if available
            const excelData = await this.extractExcelData();

            // Call API
            if (typeof APIModule !== 'undefined') {
                const response = await APIModule.sendMessage({
                    message,
                    apiKey,
                    model: model === 'auto' ? 'google/gemini-2.0-flash-lite' : model,
                    temperature,
                    cellData: this.currentSelectedCell,
                    excelData: excelData
                });

                if (response.success) {
                    this.addMessage('ai', response.content);
                    
                    // Auto-execute Excel tasks if detected
                    if (typeof Office !== 'undefined' && typeof ExcelExecutor !== 'undefined') {
                        try {
                            const execResult = await ExcelExecutor.executeFromResponse(response.content, message);
                            if (execResult.executed) {
                                this.addMessage('system', `✅ ${execResult.message}`);
                            }
                        } catch (execError) {
                            console.error('Auto-execution error:', execError);
                        }
                    }
                    
                    // Save to history
                    if (typeof StorageModule !== 'undefined') {
                        const history = StorageModule.getHistory();
                        history.push({
                            query: message,
                            response: response.content,
                            timestamp: new Date().toISOString()
                        });
                        StorageModule.setHistory(history);
                    }
                } else {
                    this.addMessage('ai', `Error: ${response.error}`);
                }
            } else {
                this.addMessage('ai', 'API module not loaded. This is a demo response.');
            }
        } catch (error) {
            console.error('Send message error:', error);
            this.addMessage('ai', `Error: ${error.message}`);
        } finally {
            this.setProcessing(false);
        }
    },

    /**
     * Add a message to the chat
     */
    addMessage(type, content) {
        const messageArea = document.getElementById('messageArea');
        if (!messageArea) return;

        // Hide welcome section
        const welcomeSection = messageArea.querySelector('.welcome-section');
        if (welcomeSection) {
            welcomeSection.classList.add('hidden');
        }

        // Hide suggestions
        const suggestionsArea = document.getElementById('suggestionsArea');
        if (suggestionsArea) {
            suggestionsArea.classList.add('hidden');
        }

        // Create message element
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}`;
        
        const messageContent = document.createElement('div');
        messageContent.className = 'message-content';
        messageContent.innerHTML = this.formatMessage(content);
        
        messageDiv.appendChild(messageContent);
        messageArea.appendChild(messageDiv);

        // Scroll to bottom
        messageArea.scrollTop = messageArea.scrollHeight;

        // Save to history
        this.messageHistory.push({ type, content, timestamp: Date.now() });
    },

    /**
     * Format message content with markdown-like formatting
     */
    formatMessage(content) {
        // Escape HTML to prevent XSS
        let formatted = content
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');

        // Convert markdown-style formatting
        // Bold: **text** or __text__
        formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
        formatted = formatted.replace(/__(.+?)__/g, '<strong>$1</strong>');
        
        // Italic: *text* or _text_
        formatted = formatted.replace(/\*(.+?)\*/g, '<em>$1</em>');
        formatted = formatted.replace(/_(.+?)_/g, '<em>$1</em>');
        
        // Code: `code`
        formatted = formatted.replace(/`(.+?)`/g, '<code>$1</code>');
        
        // Line breaks
        formatted = formatted.replace(/\n/g, '<br>');
        
        // Numbered lists: 1. item
        formatted = formatted.replace(/(\d+)\.\s+(.+?)(<br>|$)/g, '<div class="list-item">$1. $2</div>');
        
        return formatted;
    },

    showFormulaAssist(formulaText) {
        const container = document.getElementById('formulaAssist');
        const formulaSpan = document.getElementById('formulaAssistFormula');
        if (!container || !formulaSpan) return;

        formulaSpan.textContent = formulaText;
        container.classList.remove('hidden');
    },

    hideFormulaAssist() {
        const container = document.getElementById('formulaAssist');
        const formulaSpan = document.getElementById('formulaAssistFormula');
        if (container) container.classList.add('hidden');
        if (formulaSpan) formulaSpan.textContent = '';
        this.currentFormulaSuggestion = null;
    },

    async applyFormulaSuggestion() {
        if (!this.currentFormulaSuggestion || typeof ExcelExecutor === 'undefined' || typeof Office === 'undefined') {
            this.hideFormulaAssist();
            return;
        }

        const { address, formula } = this.currentFormulaSuggestion;
        this.isApplyingFormula = true;
        try {
            await ExcelExecutor.applyFormulaDirect(address, formula);
            this.hideFormulaAssist();
        } catch (error) {
            console.error('Apply formula error:', error);
        } finally {
            this.isApplyingFormula = false;
        }
    },

    /**
     * Handle suggestion card click
     */
    handleSuggestion(action) {
        console.log('Suggestion clicked:', action);
        
        const suggestions = {
            summarize: 'Summarize the selected data',
            extract: 'Extract key information from the data',
            analyze: 'Analyze the data and provide insights',
            format: 'Format and transform the data'
        };

        const messageInput = document.getElementById('messageInput');
        if (messageInput && suggestions[action]) {
            messageInput.value = suggestions[action];
            messageInput.focus();
        }
    },

    /**
     * Set processing state
     */
    setProcessing(isProcessing) {
        this.isProcessing = isProcessing;
        
        const statusIndicator = document.getElementById('statusIndicator');
        const messageInput = document.getElementById('messageInput');
        const sendButton = document.getElementById('sendButton');
        
        if (isProcessing) {
            if (statusIndicator) {
                statusIndicator.className = 'status-badge status-processing';
                statusIndicator.textContent = 'Thinking...';
            }
            if (messageInput) messageInput.disabled = true;
            if (sendButton) sendButton.disabled = true;
        } else {
            if (statusIndicator) {
                statusIndicator.className = 'status-badge status-ready';
                statusIndicator.textContent = 'Ready';
            }
            if (messageInput) messageInput.disabled = false;
            if (sendButton) sendButton.disabled = false;
        }
    },

    /**
     * Toggle dark mode
     */
    toggleDarkMode() {
        document.body.classList.toggle('dark-mode');
        const isDark = document.body.classList.contains('dark-mode');
        if (typeof StorageModule !== 'undefined') {
            StorageModule.setSetting('darkMode', isDark);
        }
    },

    /**
     * Save settings
     */
    saveSettings() {
        if (typeof StorageModule === 'undefined') {
            console.error('Storage module not available');
            return;
        }

        const apiKey = document.getElementById('apiKey').value;
        const model = document.getElementById('modelSelect').value;
        const temperature = parseFloat(document.getElementById('tempSlider').value);
        const autoSwitch = document.getElementById('autoSwitchCheckbox').checked;
        const cacheResponses = document.getElementById('cacheResponsesCheckbox').checked;
        const inlineAssist = document.getElementById('inlineAssistToggle').checked;

        StorageModule.setApiKey(apiKey);
        StorageModule.setModel(model);
        StorageModule.setSetting('temperature', temperature);
        StorageModule.setSetting('autoSwitch', autoSwitch);
        StorageModule.setSetting('cacheResponses', cacheResponses);
        StorageModule.setSetting('inlineAssist', inlineAssist);
        this.inlineAssistEnabled = inlineAssist;

        // Show success in status indicator
        const statusIndicator = document.getElementById('statusIndicator');
        if (statusIndicator) {
            statusIndicator.textContent = '✓ Saved';
            statusIndicator.className = 'status-badge status-success';
            setTimeout(() => {
                statusIndicator.textContent = 'Ready';
                statusIndicator.className = 'status-badge status-ready';
            }, 2000);
        }
        console.log('Settings saved successfully');
    },

    /**
     * Reset settings
     */
    resetSettings() {
        // Skip confirm (not supported in Office add-ins)
        if (typeof StorageModule !== 'undefined') {
            StorageModule.clearAll();
        }
        console.log('Settings reset, reloading...');
        location.reload();
    },

    /**
     * Load history
     */
    loadHistory() {
        const historyList = document.getElementById('historyList');
        if (!historyList) return;

        if (typeof StorageModule === 'undefined') {
            historyList.innerHTML = '<div class="history-empty"><i class="fas fa-inbox"></i><p>Storage not available</p></div>';
            return;
        }

        const history = StorageModule.getHistory();
        
        if (!history || history.length === 0) {
            historyList.innerHTML = '<div class="history-empty"><i class="fas fa-inbox"></i><p>No history yet</p></div>';
            return;
        }

        historyList.innerHTML = '';
        history.reverse().forEach((item, index) => {
            const historyItem = document.createElement('div');
            historyItem.className = 'history-item';
            historyItem.innerHTML = `
                <div class="history-item-text">${item.query}</div>
                <div class="history-item-time">${new Date(item.timestamp).toLocaleString()}</div>
            `;
            historyItem.addEventListener('click', () => {
                const messageInput = document.getElementById('messageInput');
                if (messageInput) {
                    messageInput.value = item.query;
                    this.switchTab('assistant');
                }
            });
            historyList.appendChild(historyItem);
        });
    },

    /**
     * Search history
     */
    searchHistory(query) {
        // Implementation for history search
        console.log('Searching history:', query);
    },

    /**
     * Clear history
     */
    clearHistory() {
        if (confirm('Clear all history?')) {
            if (typeof StorageModule !== 'undefined') {
                StorageModule.clearHistory();
            }
            this.loadHistory();
        }
    }
};

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => TaskPane.init());
} else {
    TaskPane.init();
}
