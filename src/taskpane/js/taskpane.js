/**
 * CodeExcel.AI - Task Pane Main Module
 * Clean, simple initialization for both browser and Excel
 */

const TaskPane = {
    currentSelectedCell: null,
    messageHistory: [],
    isProcessing: false,

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
            alert('Please add your OpenRouter API key in Settings');
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

            // Call API
            if (typeof APIModule !== 'undefined') {
                const response = await APIModule.sendMessage({
                    message,
                    apiKey,
                    model: model === 'auto' ? 'google/gemini-2.0-flash-lite' : model,
                    temperature,
                    cellData: this.currentSelectedCell
                });

                if (response.success) {
                    this.addMessage('ai', response.content);
                    
                    // Save to history
                    if (typeof StorageModule !== 'undefined') {
                        StorageModule.addToHistory({
                            query: message,
                            response: response.content,
                            timestamp: new Date().toISOString()
                        });
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
        messageContent.textContent = content;
        
        messageDiv.appendChild(messageContent);
        messageArea.appendChild(messageDiv);

        // Scroll to bottom
        messageArea.scrollTop = messageArea.scrollHeight;

        // Save to history
        this.messageHistory.push({ type, content, timestamp: Date.now() });
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
            alert('Storage module not available');
            return;
        }

        const apiKey = document.getElementById('apiKey').value;
        const model = document.getElementById('modelSelect').value;
        const temperature = document.getElementById('tempSlider').value;
        const autoSwitch = document.getElementById('autoSwitchCheckbox').checked;
        const cacheResponses = document.getElementById('cacheResponsesCheckbox').checked;

        StorageModule.setApiKey(apiKey);
        StorageModule.setModel(model);
        StorageModule.setSetting('temperature', temperature);
        StorageModule.setSetting('autoSwitch', autoSwitch);
        StorageModule.setSetting('cacheResponses', cacheResponses);

        alert('Settings saved successfully!');
    },

    /**
     * Reset settings
     */
    resetSettings() {
        if (confirm('Reset all settings to default?')) {
            if (typeof StorageModule !== 'undefined') {
                StorageModule.clearAll();
            }
            location.reload();
        }
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
