/**
 * API Module: OpenRouter integration for AI calls
 * Handles communication with the OpenRouter API
 */

const APIModule = {
    // OpenRouter API endpoint
    API_ENDPOINT: 'https://openrouter.ai/api/v1/chat/completions',

    // Required headers for OpenRouter
    getHeaders: function(apiKey) {
        return {
            'Authorization': `Bearer ${apiKey}`,
            'HTTP-Referer': 'https://codeexcel.ai',
            'X-Title': 'CodeExcel AI Assistant',
            'Content-Type': 'application/json'
        };
    },

    /**
     * Call OpenRouter API with user input and auto-switching support
     * @param {string} userMessage - The user message to send
     * @param {string} apiKey - The OpenRouter API key
     * @param {string} model - The model to use (or 'auto' for auto-switching)
     * @param {string} systemPrompt - The system prompt
     * @param {object} options - Additional options
     * @returns {Promise<{response: string, model: string, time: number}>} The API response with metadata
     */
    callOpenRouter: async function(userMessage, apiKey, model = 'auto', systemPrompt, options = {}) {
        // Validate inputs
        if (!apiKey || apiKey.trim() === '') {
            throw new Error('Missing API Key');
        }

        if (!userMessage || userMessage.trim() === '') {
            throw new Error('No data to process');
        }

        const startTime = Date.now();
        
        // Handle auto-switching
        let selectedModel = model;
        if (model === 'auto' && typeof ModelsManager !== 'undefined') {
            selectedModel = ModelsManager.getCurrentModel();
        }

        if (!selectedModel) {
            throw new Error('Model not specified');
        }

        const headers = this.getHeaders(apiKey);

        // Get temperature from options or use default
        const temperature = options.temperature || 0.7;

        const requestBody = {
            model: selectedModel,
            messages: [
                {
                    role: 'system',
                    content: systemPrompt || 'You are a helpful assistant for Excel data processing.'
                },
                {
                    role: 'user',
                    content: userMessage
                }
            ],
            temperature: temperature,
            max_tokens: options.maxTokens || 500,
            top_p: options.topP || 1.0,
            frequency_penalty: options.frequencyPenalty || 0,
            presence_penalty: options.presencePenalty || 0
        };

        try {
            const response = await fetch(this.API_ENDPOINT, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(requestBody)
            });

            const responseTime = Date.now() - startTime;

            // Handle rate limiting (429) with auto-switch
            if (response.status === 429) {
                if (typeof ModelsManager !== 'undefined' && model === 'auto') {
                    console.log('Rate limit exceeded, attempting to switch model...');
                    ModelsManager.recordError(selectedModel);
                    
                    // Auto-switch to next available model
                    const nextModel = ModelsManager.autoSwitch();
                    
                    if (nextModel !== selectedModel) {
                        // Retry with new model
                        return this.callOpenRouter(userMessage, apiKey, nextModel, systemPrompt, options);
                    }
                }
                throw new Error('Rate limit exceeded (429). All models are quota-limited.');
            }

            // Check for HTTP errors
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const errorMessage = errorData.error?.message || `HTTP ${response.status}`;
                
                if (typeof ModelsManager !== 'undefined') {
                    ModelsManager.recordError(selectedModel);
                }
                
                throw new Error(`API Error: ${errorMessage}`);
            }

            const data = await response.json();

            // Record successful usage
            if (typeof ModelsManager !== 'undefined') {
                ModelsManager.recordUsage(selectedModel, responseTime);
            }

            // Extract the response text
            if (data.choices && data.choices[0] && data.choices[0].message) {
                return {
                    response: data.choices[0].message.content,
                    model: selectedModel,
                    time: responseTime,
                    tokens: data.usage?.total_tokens || 0
                };
            } else {
                throw new Error('Invalid API response format');
            }
        } catch (error) {
            // Handle network errors and other issues
            if (error instanceof TypeError) {
                throw new Error('Network error: Unable to reach the API');
            }
            throw error;
        }
    },

    /**
     * Validate API key by making a test call
     * @param {string} apiKey - The API key to validate
     * @returns {Promise<boolean>} True if valid, false otherwise
     */
    validateApiKey: async function(apiKey) {
        try {
            await this.callOpenRouter(
                'Say "OK" if you received this message.',
                apiKey,
                StorageModule.getModel(),
                'You are a helpful assistant. Respond briefly.'
            );
            return true;
        } catch (error) {
            console.error('API key validation failed:', error);
            return false;
        }
    },

    /**
     * Format error message for display
     * @param {Error} error - The error object
     * @returns {string} Formatted error message
     */
    formatErrorMessage: function(error) {
        if (!error) return 'Unknown error occurred';

        if (error.message) {
            return error.message;
        }

        if (typeof error === 'string') {
            return error;
        }

        return 'An unexpected error occurred';
    }
};
