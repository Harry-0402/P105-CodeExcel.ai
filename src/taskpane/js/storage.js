/**
 * Storage Module: Secure API key and settings persistence
 * Uses browser localStorage to persist user settings across sessions
 */

const StorageModule = {
    // Storage keys
    KEYS: {
        API_KEY: 'openrouter_api_key',
        MODEL: 'openrouter_model',
        SYSTEM_PROMPT: 'openrouter_system_prompt'
    },

    // Default values
    DEFAULTS: {
        MODEL: 'google/gemini-2.0-flash-lite-preview-02-05:free',
        SYSTEM_PROMPT: 'You are an Excel data analyst. Complete the following task or analyze this data concisely.'
    },

    /**
     * Get API key from localStorage
     * @returns {string|null} The stored API key or null if not found
     */
    getApiKey: function() {
        try {
            return localStorage.getItem(this.KEYS.API_KEY);
        } catch (error) {
            console.error('Error reading API key from storage:', error);
            return null;
        }
    },

    /**
     * Set API key in localStorage
     * @param {string} apiKey - The API key to store
     * @returns {boolean} Success status
     */
    setApiKey: function(apiKey) {
        try {
            if (!apiKey || typeof apiKey !== 'string') {
                console.warn('Invalid API key format');
                return false;
            }
            localStorage.setItem(this.KEYS.API_KEY, apiKey.trim());
            return true;
        } catch (error) {
            console.error('Error saving API key to storage:', error);
            return false;
        }
    },

    /**
     * Clear API key from localStorage
     * @returns {boolean} Success status
     */
    clearApiKey: function() {
        try {
            localStorage.removeItem(this.KEYS.API_KEY);
            return true;
        } catch (error) {
            console.error('Error clearing API key:', error);
            return false;
        }
    },

    /**
     * Check if API key exists
     * @returns {boolean} True if API key is stored
     */
    hasApiKey: function() {
        return this.getApiKey() !== null && this.getApiKey() !== '';
    },

    /**
     * Get model name from localStorage
     * @returns {string} The stored model or default model
     */
    getModel: function() {
        try {
            return localStorage.getItem(this.KEYS.MODEL) || this.DEFAULTS.MODEL;
        } catch (error) {
            console.error('Error reading model from storage:', error);
            return this.DEFAULTS.MODEL;
        }
    },

    /**
     * Set model name in localStorage
     * @param {string} model - The model to store
     * @returns {boolean} Success status
     */
    setModel: function(model) {
        try {
            if (!model || typeof model !== 'string') {
                return false;
            }
            localStorage.setItem(this.KEYS.MODEL, model.trim());
            return true;
        } catch (error) {
            console.error('Error saving model to storage:', error);
            return false;
        }
    },

    /**
     * Get system prompt from localStorage
     * @returns {string} The stored prompt or default prompt
     */
    getSystemPrompt: function() {
        try {
            return localStorage.getItem(this.KEYS.SYSTEM_PROMPT) || this.DEFAULTS.SYSTEM_PROMPT;
        } catch (error) {
            console.error('Error reading system prompt from storage:', error);
            return this.DEFAULTS.SYSTEM_PROMPT;
        }
    },

    /**
     * Set system prompt in localStorage
     * @param {string} prompt - The prompt to store
     * @returns {boolean} Success status
     */
    setSystemPrompt: function(prompt) {
        try {
            if (!prompt || typeof prompt !== 'string') {
                return false;
            }
            localStorage.setItem(this.KEYS.SYSTEM_PROMPT, prompt.trim());
            return true;
        } catch (error) {
            console.error('Error saving system prompt to storage:', error);
            return false;
        }
    },

    /**
     * Clear all stored settings
     * @returns {boolean} Success status
     */
    clearAll: function() {
        try {
            localStorage.removeItem(this.KEYS.API_KEY);
            localStorage.removeItem(this.KEYS.MODEL);
            localStorage.removeItem(this.KEYS.SYSTEM_PROMPT);
            return true;
        } catch (error) {
            console.error('Error clearing all storage:', error);
            return false;
        }
    }
};
