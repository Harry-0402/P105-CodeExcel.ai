/**
 * ModelsManager.js
 * Manages multiple free OpenRouter AI models with intelligent auto-switching
 * Features: quota tracking, fallback logic, usage statistics
 */

const ModelsManager = {
    // Available free models on OpenRouter
    models: {
        'deepseek/deepseek-chat': {
            name: 'DeepSeek Chat',
            provider: 'DeepSeek',
            quotaPerMin: 30,
            costPerRequest: 0,
            category: 'general',
            priority: 1
        },
        'deepseek/deepseek-coder': {
            name: 'DeepSeek Coder',
            provider: 'DeepSeek',
            quotaPerMin: 20,
            costPerRequest: 0,
            category: 'code',
            priority: 2
        },
        'google/gemini-flash-1.5-8b': {
            name: 'Gemini Flash 1.5 8B',
            provider: 'Google',
            quotaPerMin: 10,
            costPerRequest: 0,
            category: 'general',
            priority: 3
        },
        'google/gemini-pro-1.5': {
            name: 'Gemini Pro 1.5',
            provider: 'Google',
            quotaPerMin: 5,
            costPerRequest: 0,
            category: 'advanced',
            priority: 4
        },
        'meta-llama/llama-3.1-8b-instruct:free': {
            name: 'Llama 3.1 8B',
            provider: 'Meta',
            quotaPerMin: 15,
            costPerRequest: 0,
            category: 'general',
            priority: 5
        },
        'meta-llama/llama-3.1-70b-instruct:free': {
            name: 'Llama 3.1 70B',
            provider: 'Meta',
            quotaPerMin: 8,
            costPerRequest: 0,
            category: 'advanced',
            priority: 6
        },
        'meta-llama/llama-2-70b-chat': {
            name: 'Llama 2 70B Chat',
            provider: 'Meta',
            quotaPerMin: 10,
            costPerRequest: 0,
            category: 'general',
            priority: 7
        },
        'mistralai/mistral-7b-instruct:free': {
            name: 'Mistral 7B',
            provider: 'Mistral',
            quotaPerMin: 10,
            costPerRequest: 0,
            category: 'general',
            priority: 8
        },
        'nousresearch/hermes-3-llama-3.1-405b:free': {
            name: 'Hermes 3 405B',
            provider: 'Nous Research',
            quotaPerMin: 8,
            costPerRequest: 0,
            category: 'advanced',
            priority: 9
        },
        'openchat/openchat-7b:free': {
            name: 'OpenChat 7B',
            provider: 'OpenChat',
            quotaPerMin: 25,
            costPerRequest: 0,
            category: 'general',
            priority: 10
        },
        'qwen/qwen-2-7b-instruct:free': {
            name: 'Qwen 2 7B',
            provider: 'Qwen',
            quotaPerMin: 12,
            costPerRequest: 0,
            category: 'general',
            priority: 11
        }
    },

    // Current state
    currentModel: 'deepseek/deepseek-chat',
    usageStats: {},
    lastResetTime: Date.now(),
    lastSwitchTime: 0,
    failedModels: {},

    /**
     * Initialize the models manager
     */
    init() {
        // Load saved usage stats from storage
        const saved = StorageModule.getSetting('modelsUsageStats');
        if (saved) {
            this.usageStats = saved;
        } else {
            this.resetUsageStats();
        }

        // Load saved model preferences
        const savedModel = StorageModule.getSetting('selectedModel');
        if (savedModel && this.models[savedModel]) {
            this.currentModel = savedModel;
        }

        console.log('ModelsManager initialized with model:', this.currentModel);
    },

    /**
     * Get the current model
     */
    getCurrentModel() {
        return this.currentModel;
    },

    /**
     * Get model info
     */
    getModelInfo(modelId) {
        return this.models[modelId] || null;
    },

    /**
     * Check if model has exceeded quota
     */
    isQuotaExceeded(modelId) {
        const stats = this.usageStats[modelId];
        if (!stats) return false;

        const minutesSinceReset = (Date.now() - stats.windowStart) / (1000 * 60);
        if (minutesSinceReset >= 1) {
            // Reset the count if more than 1 minute has passed
            stats.count = 0;
            stats.windowStart = Date.now();
            this.saveUsageStats();
            return false;
        }

        const model = this.models[modelId];
        return stats.count >= model.quotaPerMin;
    },

    /**
     * Record a successful request
     */
    recordUsage(modelId, responseTime = 0) {
        if (!this.usageStats[modelId]) {
            this.usageStats[modelId] = {
                count: 0,
                windowStart: Date.now(),
                totalRequests: 0,
                totalTime: 0,
                lastUsed: null,
                errors: 0
            };
        }

        const stats = this.usageStats[modelId];
        stats.count++;
        stats.totalRequests++;
        stats.totalTime += responseTime;
        stats.lastUsed = new Date().toISOString();

        this.saveUsageStats();
        this.updateUI();
    },

    /**
     * Record a failed request
     */
    recordError(modelId) {
        if (!this.usageStats[modelId]) {
            this.usageStats[modelId] = {
                count: 0,
                windowStart: Date.now(),
                totalRequests: 0,
                totalTime: 0,
                lastUsed: null,
                errors: 0
            };
        }

        this.usageStats[modelId].errors++;
        this.usageStats[modelId].lastUsed = new Date().toISOString();

        if (!this.failedModels[modelId]) {
            this.failedModels[modelId] = {
                count: 0,
                firstFailTime: Date.now()
            };
        }
        this.failedModels[modelId].count++;

        this.saveUsageStats();
    },

    /**
     * Auto-switch to next available model
     */
    autoSwitch(category = 'general', preferProvider = null) {
        const availableModels = this.getAvailableModels(category, preferProvider);

        if (availableModels.length === 0) {
            console.warn('No available models found, waiting for quota reset...');
            return this.currentModel; // Return current model, might trigger rate limit
        }

        const nextModel = availableModels[0];
        const previousModel = this.currentModel;
        this.currentModel = nextModel.id;
        this.lastSwitchTime = Date.now();

        console.log(`Auto-switching from ${previousModel} to ${this.currentModel}`);
        this.updateUI();

        // Save preference
        StorageModule.setSetting('selectedModel', this.currentModel);

        return this.currentModel;
    },

    /**
     * Get available models sorted by priority
     */
    getAvailableModels(category = null, preferProvider = null) {
        const candidates = Object.entries(this.models)
            .filter(([id, model]) => {
                if (this.isQuotaExceeded(id)) return false;
                if (this.failedModels[id] && this.failedModels[id].count > 2) return false;
                if (category && model.category !== category && category !== 'general') {
                    return model.category === 'general'; // Fallback to general
                }
                return true;
            })
            .map(([id, model]) => ({
                id,
                ...model,
                score: this.calculateScore(id, preferProvider)
            }))
            .sort((a, b) => a.score - b.score);

        return candidates;
    },

    /**
     * Calculate priority score for a model
     */
    calculateScore(modelId, preferProvider = null) {
        const model = this.models[modelId];
        let score = model.priority;

        // Prefer models from the same provider if specified
        if (preferProvider && model.provider === preferProvider) {
            score -= 100; // Strongly prefer same provider
        }

        // Prefer models with higher quota
        score -= (model.quotaPerMin / 100);

        // Penalize models that have been used recently
        const stats = this.usageStats[modelId];
        if (stats && stats.lastUsed) {
            const minutesSinceUse = (Date.now() - new Date(stats.lastUsed)) / (1000 * 60);
            if (minutesSinceUse < 5) {
                score += 50; // Penalize recent usage
            }
        }

        return score;
    },

    /**
     * Get usage statistics for all models
     */
    getStats() {
        const stats = {
            currentModel: this.currentModel,
            totalRequests: 0,
            totalModelsUsed: 0,
            models: {}
        };

        for (const [modelId, usage] of Object.entries(this.usageStats)) {
            if (usage.totalRequests > 0) {
                stats.models[modelId] = {
                    name: this.models[modelId]?.name,
                    requests: usage.totalRequests,
                    errors: usage.errors,
                    avgTime: usage.totalTime / usage.totalRequests,
                    lastUsed: usage.lastUsed
                };
                stats.totalRequests += usage.totalRequests;
                stats.totalModelsUsed++;
            }
        }

        return stats;
    },

    /**
     * Reset usage statistics
     */
    resetUsageStats() {
        this.usageStats = {};
        for (const modelId of Object.keys(this.models)) {
            this.usageStats[modelId] = {
                count: 0,
                windowStart: Date.now(),
                totalRequests: 0,
                totalTime: 0,
                lastUsed: null,
                errors: 0
            };
        }
        this.lastResetTime = Date.now();
        this.saveUsageStats();
    },

    /**
     * Save usage stats to storage
     */
    saveUsageStats() {
        StorageModule.setSetting('modelsUsageStats', this.usageStats);
    },

    /**
     * Update UI with current model and stats
     */
    updateUI() {
        const modelBadge = document.getElementById('currentModel');
        if (modelBadge) {
            const model = this.models[this.currentModel];
            modelBadge.textContent = model?.name?.split(' ')[0] || 'Auto';
        }

        // Update stats section if visible
        const statsGrid = document.getElementById('modelStats');
        if (statsGrid) {
            this.renderStats(statsGrid);
        }
    },

    /**
     * Render statistics UI
     */
    renderStats(container) {
        const stats = this.getStats();
        container.innerHTML = '';

        if (stats.totalRequests === 0) {
            container.innerHTML = '<div style="grid-column: 1/-1; text-align: center; padding: 16px; color: var(--color-text-secondary);">No statistics yet</div>';
            return;
        }

        for (const [modelId, data] of Object.entries(stats.models)) {
            const card = document.createElement('div');
            card.className = 'stat-card';
            card.innerHTML = `
                <div class="stat-label">${data.name}</div>
                <div class="stat-value">${data.requests}</div>
                <div class="stat-unit">requests</div>
            `;
            container.appendChild(card);
        }
    },

    /**
     * Set model manually
     */
    setModel(modelId) {
        if (!this.models[modelId]) {
            console.error('Invalid model ID:', modelId);
            return false;
        }
        this.currentModel = modelId;
        StorageModule.setSetting('selectedModel', modelId);
        this.updateUI();
        return true;
    },

    /**
     * Get all models list
     */
    getAllModels() {
        return Object.entries(this.models).map(([id, model]) => ({
            id,
            ...model
        }));
    },

    /**
     * Check if should auto-switch (convenience method)
     */
    shouldAutoSwitch() {
        return this.isQuotaExceeded(this.currentModel);
    }
};

// Initialize when document is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => ModelsManager.init());
} else {
    ModelsManager.init();
}
