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
     * Call OpenRouter API with user input
     * @param {string} userMessage - The user message to send
     * @param {string} apiKey - The OpenRouter API key
     * @param {string} model - The model to use
     * @param {string} systemPrompt - The system prompt
     * @returns {Promise<string>} The API response text
     */
    callOpenRouter: async function(userMessage, apiKey, model, systemPrompt) {
        // Validate inputs
        if (!apiKey || apiKey.trim() === '') {
            throw new Error('Missing API Key');
        }

        if (!userMessage || userMessage.trim() === '') {
            throw new Error('No data to process');
        }

        if (!model) {
            throw new Error('Model not specified');
        }

        const headers = this.getHeaders(apiKey);

        const requestBody = {
            model: model,
            messages: [
                {
                    role: 'system',
                    content: systemPrompt
                },
                {
                    role: 'user',
                    content: userMessage
                }
            ],
            temperature: 0.7,
            max_tokens: 500,
            top_p: 1.0,
            frequency_penalty: 0,
            presence_penalty: 0
        };

        try {
            const response = await fetch(this.API_ENDPOINT, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(requestBody)
            });

            // Check for HTTP errors
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const errorMessage = errorData.error?.message || `HTTP ${response.status}`;
                throw new Error(`API Error: ${errorMessage}`);
            }

            const data = await response.json();

            // Extract the response text
            if (data.choices && data.choices[0] && data.choices[0].message) {
                return data.choices[0].message.content;
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
