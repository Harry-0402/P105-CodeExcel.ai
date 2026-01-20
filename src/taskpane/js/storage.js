/**
 * Storage Module: Secure API key and settings persistence
 * Uses browser localStorage to persist user settings across sessions
 */

const StorageModule = {
    // Storage keys
    KEYS: {
        API_KEY: 'openrouter_api_key',
        MODEL: 'openrouter_model',
        SYSTEM_PROMPT: 'openrouter_system_prompt',
        HISTORY: 'message_history',
        CACHE: 'response_cache',
        SETTINGS: 'app_settings'
    },

    // Default values
    DEFAULTS: {
        MODEL: 'deepseek/deepseek-chat',
        SYSTEM_PROMPT: 'You are an Excel AI assistant that automatically executes tasks. When user asks to create tables, insert data, format cells, or perform calculations, respond with clear confirmation. The system will automatically execute the task in Excel.',
        TEMPERATURE: 0.7,
        AUTO_SWITCH: true,
        CACHE_RESPONSES: true,
        DARK_MODE: false,
        INLINE_ASSIST: true
    },

    /**
     * Initialize storage module
     */
    init() {
        console.log('StorageModule initialized');
        // Load dark mode preference if saved
        if (this.getSetting('darkMode')) {
            document.body.classList.add('dark-mode');
        }
    },

    /**
     * Get API key from localStorage
     * @returns {string|null} The stored API key or null if not found
     */
    getApiKey() {
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
    setApiKey(apiKey) {
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
    clearApiKey() {
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
    hasApiKey() {
        return this.getApiKey() !== null && this.getApiKey() !== '';
    },

    /**
     * Get model name from localStorage
     * @returns {string} The stored model or default model
     */
    getModel() {
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
     * Get message history
     * @returns {array} Message history array
     */
    getHistory() {
        try {
            const data = localStorage.getItem(this.KEYS.HISTORY);
            return data ? JSON.parse(data) : [];
        } catch (error) {
            console.error('Error reading history:', error);
            return [];
        }
    },

    /**
     * Set message history
     * @param {array} history - History array to store
     */
    setHistory(history) {
        try {
            localStorage.setItem(this.KEYS.HISTORY, JSON.stringify(history));
            return true;
        } catch (error) {
            console.error('Error saving history:', error);
            return false;
        }
    },

    /**
     * Get a specific setting
     * @param {string} key - The setting key
     * @returns {*} The setting value
     */
    getSetting(key) {
        try {
            const settings = localStorage.getItem(this.KEYS.SETTINGS);
            const parsed = settings ? JSON.parse(settings) : {};
            return parsed[key] !== undefined ? parsed[key] : this.DEFAULTS[key.toUpperCase()];
        } catch (error) {
            console.error('Error reading setting:', error);
            return null;
        }
    },

    /**
     * Set a specific setting
     * @param {string} key - The setting key
     * @param {*} value - The setting value
     */
    setSetting(key, value) {
        try {
            const settings = localStorage.getItem(this.KEYS.SETTINGS);
            const parsed = settings ? JSON.parse(settings) : {};
            parsed[key] = value;
            localStorage.setItem(this.KEYS.SETTINGS, JSON.stringify(parsed));
            return true;
        } catch (error) {
            console.error('Error saving setting:', error);
            return false;
        }
    },

    /**
     * Get all settings
     * @returns {object} All settings object
     */
    getAllSettings() {
        try {
            const settings = localStorage.getItem(this.KEYS.SETTINGS);
            const parsed = settings ? JSON.parse(settings) : {};
            return {
                apiKey: this.getApiKey() || '',
                model: parsed.model || this.getModel() || this.DEFAULTS.MODEL,
                temperature: parsed.temperature ?? this.DEFAULTS.TEMPERATURE,
                autoSwitch: parsed.autoSwitch ?? this.DEFAULTS.AUTO_SWITCH,
                cacheResponses: parsed.cacheResponses ?? this.DEFAULTS.CACHE_RESPONSES,
                darkMode: parsed.darkMode ?? this.DEFAULTS.DARK_MODE,
                inlineAssist: parsed.inlineAssist ?? this.DEFAULTS.INLINE_ASSIST,
                systemPrompt: parsed.systemPrompt || this.getSystemPrompt()
            };
        } catch (error) {
            console.error('Error reading all settings:', error);
            return {
                apiKey: this.getApiKey() || '',
                model: this.DEFAULTS.MODEL,
                temperature: this.DEFAULTS.TEMPERATURE,
                autoSwitch: this.DEFAULTS.AUTO_SWITCH,
                cacheResponses: this.DEFAULTS.CACHE_RESPONSES,
                darkMode: this.DEFAULTS.DARK_MODE,
                inlineAssist: this.DEFAULTS.INLINE_ASSIST,
                systemPrompt: this.DEFAULTS.SYSTEM_PROMPT
            };
        }
    },

    /**
     * Clear all stored settings and cached values
     */
    clearAll() {
        try {
            localStorage.removeItem(this.KEYS.API_KEY);
            localStorage.removeItem(this.KEYS.MODEL);
            localStorage.removeItem(this.KEYS.SYSTEM_PROMPT);
            localStorage.removeItem(this.KEYS.HISTORY);
            localStorage.removeItem(this.KEYS.CACHE);
            localStorage.removeItem(this.KEYS.SETTINGS);
            return true;
        } catch (error) {
            console.error('Error clearing all storage:', error);
            return false;
        }
    },

    /**
     * Cache a response
     * @param {string} key - The cache key (usually the input hash)
     * @param {string} value - The cached response
     */
    cacheResponse(key, value) {
        try {
            const cache = localStorage.getItem(this.KEYS.CACHE);
            const parsed = cache ? JSON.parse(cache) : {};
            parsed[key] = {
                value,
                timestamp: Date.now()
            };
            localStorage.setItem(this.KEYS.CACHE, JSON.stringify(parsed));
            return true;
        } catch (error) {
            console.error('Error caching response:', error);
            return false;
        }
    },

    /**
     * Get cached response
     * @param {string} key - The cache key
     * @returns {string|null} The cached response or null
     */
    getCachedResponse(key) {
        try {
            const cache = localStorage.getItem(this.KEYS.CACHE);
            if (!cache) return null;
            const parsed = JSON.parse(cache);
            const cached = parsed[key];
            if (!cached) return null;
            
            // Invalidate cache after 24 hours
            const ageMs = Date.now() - cached.timestamp;
            const age24h = 24 * 60 * 60 * 1000;
            if (ageMs > age24h) {
                delete parsed[key];
                localStorage.setItem(this.KEYS.CACHE, JSON.stringify(parsed));
                return null;
            }
            
            return cached.value;
        } catch (error) {
            console.error('Error reading cached response:', error);
            return null;
        }
    },

    /**
     * Clear all settings
     */
    clearAllSettings() {
        try {
            localStorage.clear();
            return true;
        } catch (error) {
            console.error('Error clearing settings:', error);
            return false;
        }
    }
};
