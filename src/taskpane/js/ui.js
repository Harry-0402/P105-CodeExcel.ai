/**
 * UI Module: Status indicators and user interface state management
 * Handles visual feedback for user actions and system state
 */

const UIModule = {
    // Status types
    STATUS: {
        IDLE: 'idle',
        PROCESSING: 'processing',
        SUCCESS: 'success',
        ERROR: 'error'
    },

    // DOM elements cache
    elements: {
        status: null,
        statusText: null,
        response: null,
        processButton: null,
        apiKeyInput: null,
        modelInput: null,
        systemPromptInput: null,
        saveButton: null,
        clearButton: null,
        settingsStatus: null
    },

    /**
     * Initialize UI module and cache DOM elements
     */
    init: function() {
        this.elements.status = document.getElementById('status');
        this.elements.statusText = document.getElementById('statusText');
        this.elements.response = document.getElementById('response');
        this.elements.processButton = document.getElementById('processButton');
        this.elements.apiKeyInput = document.getElementById('apiKeyInput');
        this.elements.modelInput = document.getElementById('modelInput');
        this.elements.systemPromptInput = document.getElementById('systemPromptInput');
        this.elements.saveButton = document.getElementById('saveButton');
        this.elements.clearButton = document.getElementById('clearButton');
        this.elements.settingsStatus = document.getElementById('settingsStatus');
    },

    /**
     * Update status indicator
     * @param {string} status - Status type (idle, processing, success, error)
     * @param {string} message - Status message to display
     */
    setStatus: function(status, message) {
        if (!this.elements.status) return;

        // Remove all status classes
        this.elements.status.classList.remove(
            'status-idle',
            'status-processing',
            'status-success',
            'status-error'
        );

        // Add appropriate class and update text
        this.elements.status.classList.add(`status-${status}`);
        this.elements.statusText.textContent = message;
    },

    /**
     * Set idle status
     * @param {string} message - Optional message
     */
    setIdle: function(message = 'Ready') {
        this.setStatus(this.STATUS.IDLE, message);
    },

    /**
     * Set processing status
     * @param {string} message - Optional message
     */
    setProcessing: function(message = 'Processing...') {
        this.setStatus(this.STATUS.PROCESSING, message);
    },

    /**
     * Set success status
     * @param {string} message - Optional message
     */
    setSuccess: function(message = 'Success') {
        this.setStatus(this.STATUS.SUCCESS, message);
        setTimeout(() => this.setIdle(), 3000);
    },

    /**
     * Set error status
     * @param {string} message - Optional message
     */
    setError: function(message = 'Error') {
        this.setStatus(this.STATUS.ERROR, message);
    },

    /**
     * Display AI response in response box
     * @param {string} response - The AI response text
     */
    displayResponse: function(response) {
        if (!this.elements.response) return;

        this.elements.response.innerHTML = `<pre class="response-content">${this.escapeHtml(response)}</pre>`;
    },

    /**
     * Clear response box
     */
    clearResponse: function() {
        if (this.elements.response) {
            this.elements.response.innerHTML = '<p class="response-placeholder">AI response will appear here</p>';
        }
    },

    /**
     * Enable/disable process button
     * @param {boolean} enabled - Whether to enable the button
     */
    setProcessButtonEnabled: function(enabled) {
        if (this.elements.processButton) {
            this.elements.processButton.disabled = !enabled;
        }
    },

    /**
     * Load settings into form inputs
     * @param {object} settings - Settings object with apiKey, model, systemPrompt
     */
    loadSettingsForm: function(settings) {
        if (this.elements.apiKeyInput) {
            this.elements.apiKeyInput.value = settings.apiKey || '';
        }
        if (this.elements.modelInput) {
            this.elements.modelInput.value = settings.model || StorageModule.DEFAULTS.MODEL;
        }
        if (this.elements.systemPromptInput) {
            this.elements.systemPromptInput.value = settings.systemPrompt || StorageModule.DEFAULTS.SYSTEM_PROMPT;
        }
    },

    /**
     * Get settings from form inputs
     * @returns {object} Settings object
     */
    getSettingsForm: function() {
        return {
            apiKey: this.elements.apiKeyInput ? this.elements.apiKeyInput.value : '',
            model: this.elements.modelInput ? this.elements.modelInput.value : '',
            systemPrompt: this.elements.systemPromptInput ? this.elements.systemPromptInput.value : ''
        };
    },

    /**
     * Show settings save status message
     * @param {string} type - 'success' or 'error'
     * @param {string} message - Status message
     */
    showSettingsStatus: function(type, message) {
        if (!this.elements.settingsStatus) return;

        this.elements.settingsStatus.className = `settings-status ${type}`;
        this.elements.settingsStatus.textContent = message;

        if (type === 'success') {
            setTimeout(() => {
                this.elements.settingsStatus.textContent = '';
                this.elements.settingsStatus.className = 'settings-status';
            }, 3000);
        }
    },

    /**
     * Escape HTML special characters
     * @param {string} text - Text to escape
     * @returns {string} Escaped text
     */
    escapeHtml: function(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    /**
     * Update selected cell display
     * @param {string} cellAddress - The cell address (e.g., "A1")
     */
    updateSelectedCell: function(cellAddress) {
        const elem = document.getElementById('selectedCell');
        if (elem) {
            elem.textContent = cellAddress || 'None';
        }
    }
};
