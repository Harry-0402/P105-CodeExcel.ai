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
            // Use the deployed origin for OpenRouter attribution
            'HTTP-Referer': 'https://harry-0402.github.io/P105-CodeExcel.ai',
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
                let errorMessage = `HTTP ${response.status}`;
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.error?.message || errorData.message || errorMessage;
                    console.error('API Error Response:', errorData);
                } catch (parseError) {
                    console.error('Could not parse error response:', parseError);
                    const text = await response.text().catch(() => '');
                    console.error('Raw error response:', text);
                }
                
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
            console.error('API call error:', error);
            // Re-throw API errors as-is (they have useful messages)
            if (error.message && error.message.startsWith('API Error:')) {
                throw error;
            }
            // Handle fetch/network errors
            if (error instanceof TypeError || error.name === 'TypeError') {
                throw new Error('Network error: Unable to reach OpenRouter. Check your internet connection or firewall settings.');
            }
            // For other errors, include the original message
            throw new Error(`Request failed: ${error.message || error}`);
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
    },

    /**
     * High-level sendMessage used by taskpane
     */
    async sendMessage({ message, apiKey, model, temperature = 0.7, cellData, excelData, systemPrompt }) {
        const prompt = systemPrompt || StorageModule?.getSystemPrompt?.() || 'You are a helpful assistant for Excel data processing.';
        
        // Build context with Excel data
        let contextMessage = message;
        
        if (excelData) {
            contextMessage += '\n\n--- Excel Context ---';
            
            // Add sheet info
            if (excelData.sheetName) {
                contextMessage += `\nSheet: ${excelData.sheetName}`;
            }
            
            // Add selected range info
            if (excelData.selectedRange && excelData.selectedRange.values) {
                contextMessage += `\n\nSelected Range (${excelData.selectedRange.address}):`;
                contextMessage += `\n${this.formatDataAsTable(excelData.selectedRange.values)}`;
            }
            
            // Add table info if available
            if (excelData.tables && excelData.tables.length > 0) {
                contextMessage += '\n\nAvailable Tables:';
                excelData.tables.forEach(table => {
                    contextMessage += `\n\nTable: ${table.name} (${table.rowCount} rows)`;
                    contextMessage += `\nHeaders: ${table.headers.join(', ')}`;
                    // Include first few rows of data
                    if (table.values.length > 1) {
                        contextMessage += '\nSample Data:';
                        const sampleRows = table.values.slice(0, Math.min(6, table.values.length));
                        contextMessage += `\n${this.formatDataAsTable(sampleRows)}`;
                    }
                });
            }
            
            // Add used range info if no tables
            if ((!excelData.tables || excelData.tables.length === 0) && excelData.usedRange && excelData.usedRange.values) {
                contextMessage += `\n\nSheet Data (${excelData.usedRange.address}):`;
                contextMessage += `\n${this.formatDataAsTable(excelData.usedRange.values)}`;
                if (excelData.usedRange.truncated) {
                    contextMessage += `\n${excelData.usedRange.note}`;
                }
            }
        } else if (cellData) {
            contextMessage += `\n\nSelected cells: ${cellData}`;
        }
        
        const result = await this.callOpenRouter(contextMessage, apiKey, model, prompt, { temperature });
        return { success: true, content: result.response, model: result.model, time: result.time };
    },

    /**
     * Format 2D array as readable table
     */
    formatDataAsTable(values) {
        if (!values || values.length === 0) return '';
        
        let output = '';
        values.forEach((row, i) => {
            output += '\n' + row.map(cell => {
                if (cell === null || cell === undefined) return '';
                if (typeof cell === 'number') return cell.toString();
                return String(cell);
            }).join(' | ');
        });
        return output;
    },

    /**
     * Get a single best formula suggestion based on a cell edit
     */
    async getFormulaSuggestion({ apiKey, model = 'auto', temperature = 0.4, address, userFormula, excelData }) {
        if (!apiKey) throw new Error('Missing API Key');
        if (!address) throw new Error('Missing cell address');

        const systemPrompt = 'You are an Excel formula assistant. Given context and a partial or initial user formula, return exactly one complete Excel formula starting with =. No prose, no code fences, no explanation.';

        let contextMessage = `Active cell: ${address}`;
        if (userFormula) {
            contextMessage += `\nUser typed: ${userFormula}`;
        }

        if (excelData) {
            if (excelData.selectedRange && excelData.selectedRange.values) {
                contextMessage += `\nSelected range ${excelData.selectedRange.address}:`;
                contextMessage += `\n${this.formatDataAsTable(excelData.selectedRange.values)}`;
            }

            if (excelData.tables && excelData.tables.length > 0) {
                contextMessage += '\nTables:';
                excelData.tables.forEach((table) => {
                    contextMessage += `\n${table.name} (${table.address}) headers: ${table.headers.join(', ')}`;
                });
            }
        }

        const result = await this.callOpenRouter(contextMessage, apiKey, model, systemPrompt, {
            temperature,
            maxTokens: 120,
            topP: 0.9
        });

        // Extract first line starting with '='
        const lines = result.response.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const best = lines.find(l => l.startsWith('=')) || lines[0] || '';

        return { success: true, formula: best, raw: result.response, model: result.model };
    }
};
