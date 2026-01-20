/**
 * Excel Executor Module
 * Automatically executes Excel tasks based on AI responses
 */

const ExcelExecutor = {
    /**
     * Parse AI response and execute Excel operations
     * @param {string} aiResponse - The AI response text
     * @param {string} userQuery - The original user query
     */
    async executeFromResponse(aiResponse, userQuery) {
        try {
            // Detect intent from both user query and AI response
            let intent = this.detectIntent(userQuery);
            
            // If no intent from user query, check if AI response contains action keywords
            if (!intent) {
                console.log('Checking AI response for action keywords...');
                intent = this.detectActionInResponse(aiResponse, userQuery);
            }
            
            console.log('Detected intent:', intent);

            if (!intent) {
                console.log('No executable intent detected');
                return { executed: false, reason: 'No executable task detected' };
            }

            // Execute based on intent
            switch (intent.type) {
                case 'CREATE_TABLE':
                    return await this.createTable(intent.params);
                case 'INSERT_DATA':
                    return await this.insertData(intent.params);
                case 'REPLACE_DATA':
                    return await this.insertData(intent.params);
                case 'FORMAT_CELLS':
                    return await this.formatCells(intent.params);
                case 'CALCULATE':
                    return await this.addFormula(intent.params);
                case 'SORT_DATA':
                    return await this.sortData(intent.params);
                case 'FILTER_DATA':
                    return await this.filterData(intent.params);
                default:
                    return { executed: false, reason: 'Unknown task type' };
            }
        } catch (error) {
            console.error('Excel execution error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Detect if AI response contains action proposals
     * @param {string} aiResponse - The AI response text
     * @param {string} userQuery - Original user query for context
     * @returns {object|null} Intent object with type and params
     */
    detectActionInResponse(aiResponse, userQuery) {
        const lowerResponse = aiResponse.toLowerCase();
        
        // Look for cell ranges in the response (e.g., "A1:E6")
        const rangeMatch = aiResponse.match(/([A-Z]+\d+):([A-Z]+\d+)/);
        const cellRange = rangeMatch ? rangeMatch[0] : 'A1';
        
        // Replace/rewrite intent
        if (lowerResponse.includes('will replace') || lowerResponse.includes('will rewrite') || 
            lowerResponse.includes('replace the') || lowerResponse.includes('will add the data')) {
            return {
                type: 'REPLACE_DATA',
                params: { startCell: cellRange.split(':')[0] || 'A1', headers: [], rows: 5 }
            };
        }
        
        // Create intent
        if (lowerResponse.includes('will create') || lowerResponse.includes('create the table') ||
            lowerResponse.includes('creating a table')) {
            return {
                type: 'CREATE_TABLE',
                params: { startCell: cellRange.split(':')[0] || 'A1', headers: [], rows: 5 }
            };
        }
        
        // Insert intent
        if (lowerResponse.includes('will insert') || lowerResponse.includes('inserting')) {
            return {
                type: 'INSERT_DATA',
                params: { startCell: cellRange.split(':')[0] || 'A1', headers: [], rows: 5 }
            };
        }
        
        // Format intent
        if (lowerResponse.includes('will format') || lowerResponse.includes('formatting')) {
            return {
                type: 'FORMAT_CELLS',
                params: { range: cellRange, bold: true }
            };
        }
        
        return null;
    },

    /**
     * Detect user intent and extract parameters
     * @param {string} query - User query
     * @returns {object|null} Intent object with type and params
     */
    detectIntent(query) {
        const lowerQuery = query.toLowerCase();

        // Action-based detection (explicit user requests)
        // Create table intent
        if (lowerQuery.includes('create') && (lowerQuery.includes('table') || lowerQuery.includes('data'))) {
            return {
                type: 'CREATE_TABLE',
                params: this.parseTableRequest(query)
            };
        }

        // Insert data intent
        if (lowerQuery.includes('insert') || lowerQuery.includes('add ') || lowerQuery.includes('populate')) {
            return {
                type: 'INSERT_DATA',
                params: this.parseDataRequest(query)
            };
        }

        // Format cells intent
        if (lowerQuery.includes('format') || lowerQuery.includes('style') || lowerQuery.includes('bold') || 
            lowerQuery.includes('color') || lowerQuery.includes('highlight')) {
            return {
                type: 'FORMAT_CELLS',
                params: this.parseFormatRequest(query)
            };
        }

        // Calculate/Formula intent
        if (lowerQuery.includes('calculate') || lowerQuery.includes('formula') || lowerQuery.includes('sum') ||
            lowerQuery.includes('total') || lowerQuery.includes('count') || lowerQuery.includes('average')) {
            return {
                type: 'CALCULATE',
                params: this.parseFormulaRequest(query)
            };
        }

        // Sort intent
        if (lowerQuery.includes('sort')) {
            return {
                type: 'SORT_DATA',
                params: this.parseSortRequest(query)
            };
        }

        // Filter intent
        if (lowerQuery.includes('filter')) {
            return {
                type: 'FILTER_DATA',
                params: this.parseFilterRequest(query)
            };
        }

        return null;
    },

    /**
     * Parse table creation request
     */
    parseTableRequest(query) {
        // Extract cell reference (e.g., A1, B2)
        const cellMatch = query.match(/(?:at\s+|cell\s+|range\s+)?([A-Z]+\d+)/i);
        const startCell = cellMatch ? cellMatch[1] : 'A1';

        // Extract headers
        let headers = [];
        
        // Look for explicit headers in quotes or after "headers:" keyword
        const headerMatches = query.match(/headers?[:\s]+([^.!?]+)/i);
        if (headerMatches) {
            const headerText = headerMatches[1];
            // Match quoted strings or comma-separated values
            const matches = headerText.match(/'([^']+)'|"([^"]+)"|([A-Za-z\s]+)/g);
            if (matches) {
                matches.forEach(match => {
                    const cleaned = match.replace(/['",]/g, '').trim();
                    if (cleaned && cleaned.length > 0) {
                        headers.push(cleaned);
                    }
                });
            }
        }
        
        // If no headers found, use default smart headers based on common table types
        if (headers.length === 0) {
            // Check if user mentioned specific types
            if (query.toLowerCase().includes('product') || query.toLowerCase().includes('sales')) {
                headers = ['Date', 'Product Name', 'Category', 'Item', 'Sales Amount'];
            } else if (query.toLowerCase().includes('employee') || query.toLowerCase().includes('staff')) {
                headers = ['ID', 'Name', 'Department', 'Position', 'Salary'];
            } else if (query.toLowerCase().includes('student') || query.toLowerCase().includes('grade')) {
                headers = ['ID', 'Name', 'Subject', 'Score', 'Grade'];
            } else {
                // Default headers
                headers = ['Column 1', 'Column 2', 'Column 3', 'Column 4', 'Column 5'];
            }
        }

        // Extract number of rows
        const rowMatch = query.match(/(\d+)\s+rows?/i);
        const rows = rowMatch ? parseInt(rowMatch[1]) : 5;

        return { startCell, headers, rows };
    },

    /**
     * Parse data insertion request
     */
    parseDataRequest(query) {
        // Similar parsing logic for data insertion
        return this.parseTableRequest(query);
    },

    /**
     * Parse format request
     */
    parseFormatRequest(query) {
        return {
            range: this.extractRange(query),
            bold: query.toLowerCase().includes('bold'),
            italic: query.toLowerCase().includes('italic'),
            color: this.extractColor(query)
        };
    },

    /**
     * Parse formula request
     */
    parseFormulaRequest(query) {
        return {
            range: this.extractRange(query),
            formula: this.extractFormula(query)
        };
    },

    /**
     * Parse sort request
     */
    parseSortRequest(query) {
        return {
            range: this.extractRange(query),
            ascending: !query.toLowerCase().includes('descending')
        };
    },

    /**
     * Parse filter request
     */
    parseFilterRequest(query) {
        return {
            range: this.extractRange(query),
            criteria: this.extractFilterCriteria(query)
        };
    },

    /**
     * Extract cell range from query
     */
    extractRange(query) {
        const rangeMatch = query.match(/([A-Z]+\d+):([A-Z]+\d+)/i);
        return rangeMatch ? rangeMatch[0] : null;
    },

    /**
     * Extract color from query
     */
    extractColor(query) {
        const colors = {
            red: '#FF0000',
            blue: '#0000FF',
            green: '#00FF00',
            yellow: '#FFFF00',
            orange: '#FFA500'
        };
        for (const [name, hex] of Object.entries(colors)) {
            if (query.toLowerCase().includes(name)) {
                return hex;
            }
        }
        return null;
    },

    /**
     * Extract formula from query
     */
    extractFormula(query) {
        const formulaMatch = query.match(/=\s*[A-Z]+\([^)]+\)/i);
        return formulaMatch ? formulaMatch[0] : null;
    },

    /**
     * Extract filter criteria
     */
    extractFilterCriteria(query) {
        // Simple extraction - can be enhanced
        return query;
    },

    /**
     * Create an Excel table
     */
    async createTable({ startCell = 'A1', headers = [], rows = 5 }) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                // Parse start cell
                const cellMatch = startCell.match(/([A-Z]+)(\d+)/i);
                if (!cellMatch) {
                    throw new Error('Invalid start cell format');
                }
                
                const startCol = cellMatch[1].toUpperCase();
                const startRowNum = parseInt(cellMatch[2]);

                // Use provided headers or create default ones
                const finalHeaders = headers.length > 0 ? headers : ['Column 1', 'Column 2', 'Column 3'];
                
                // Create headers
                const lastCol = this.getColumnLetter(finalHeaders.length - 1, startCol);
                const headerRange = sheet.getRange(`${startCol}${startRowNum}:${lastCol}${startRowNum}`);
                headerRange.values = [finalHeaders];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = '#4472C4';
                headerRange.format.font.color = 'white';

                // Generate sample data
                const data = [];
                const rowCount = rows && rows > 0 ? rows : 5;
                
                for (let i = 1; i <= rowCount; i++) {
                    const row = [];
                    for (let j = 0; j < finalHeaders.length; j++) {
                        const header = finalHeaders[j].toLowerCase();
                        if (header.includes('date')) {
                            row.push(this.generateSampleDate(i));
                        } else if (header.includes('name') || header.includes('product')) {
                            row.push(`${finalHeaders[j]} ${i}`);
                        } else if (header.includes('amount') || header.includes('price') || header.includes('sales')) {
                            row.push(Math.floor(Math.random() * 1000) + 100);
                        } else if (header.includes('category')) {
                            row.push(['Electronics', 'Furniture', 'Clothing', 'Food'][i % 4]);
                        } else {
                            row.push(`Item ${i}`);
                        }
                    }
                    data.push(row);
                }

                // Insert data
                if (data.length > 0) {
                    const dataStartRow = startRowNum + 1;
                    const dataEndRow = dataStartRow + rowCount - 1;
                    const dataRange = sheet.getRange(`${startCol}${dataStartRow}:${lastCol}${dataEndRow}`);
                    dataRange.values = data;
                }

                // Create table
                const tableRange = sheet.getRange(`${startCol}${startRowNum}:${lastCol}${startRowNum + rowCount}`);
                const table = sheet.tables.add(tableRange, true);
                table.name = `Table${Date.now()}`;
                table.style = 'TableStyleMedium2';

                await context.sync();
                return { executed: true, message: `✅ Created table with ${finalHeaders.length} columns and ${rowCount} rows` };
            });
        } catch (error) {
            console.error('Create table error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Insert data into Excel
     */
    async insertData({ startCell = 'A1', headers = [], rows = 5 }) {
        return await this.createTable({ startCell, headers, rows });
    },

    /**
     * Format cells
     */
    async formatCells({ range, bold, italic, color }) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = range ? sheet.getRange(range) : sheet.getUsedRange();
                
                if (bold) targetRange.format.font.bold = true;
                if (italic) targetRange.format.font.italic = true;
                if (color) targetRange.format.fill.color = color;

                await context.sync();
                return { executed: true, message: '✅ Formatting applied' };
            });
        } catch (error) {
            console.error('Format cells error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Add formula
     */
    async addFormula({ range, formula }) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(range);
                targetRange.formulas = [[formula]];

                await context.sync();
                return { executed: true, message: '✅ Formula added' };
            });
        } catch (error) {
            console.error('Add formula error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Sort data
     */
    async sortData({ range, ascending }) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(range);
                targetRange.sort.apply([{ key: 0, ascending: ascending !== false }]);

                await context.sync();
                return { executed: true, message: '✅ Data sorted' };
            });
        } catch (error) {
            console.error('Sort data error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Filter data
     */
    async filterData({ range, criteria }) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(range);
                targetRange.autoFilter.apply(targetRange);

                await context.sync();
                return { executed: true, message: '✅ Filter applied' };
            });
        } catch (error) {
            console.error('Filter data error:', error);
            return { executed: false, error: error.message };
        }
    },

    /**
     * Helper: Generate sample date
     */
    generateSampleDate(offset) {
        const date = new Date();
        date.setDate(date.getDate() - offset);
        return date.toISOString().split('T')[0];
    },

    /**
     * Helper: Get column letter from index
     */
    getColumnLetter(index, startCol = 'A') {
        const startIndex = startCol.charCodeAt(0) - 65;
        const targetIndex = startIndex + index;
        return String.fromCharCode(65 + targetIndex);
    }
};
