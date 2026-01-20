/**
 * Task Pane Module: Main application logic
 * Coordinates tab navigation, button handlers, and Excel integration
 */

let currentSelectedRange = null;

/**
 * Initialize the application
 */
async function initializeApp() {
    // Initialize UI module and cache DOM elements
    UIModule.init();

    // Initialize Office.js
    await Office.onReady(async (info) => {
        if (info.host !== Office.HostType.Excel) {
            UIModule.setError('This add-in only works with Excel');
            return;
        }

        // Load stored settings into form
        const settings = {
            apiKey: StorageModule.getApiKey(),
            model: StorageModule.getModel(),
            systemPrompt: StorageModule.getSystemPrompt()
        };
        UIModule.loadSettingsForm(settings);

        // Check if API key exists
        if (!StorageModule.hasApiKey()) {
            UIModule.setError('API Key required - visit Settings tab');
            UIModule.setProcessButtonEnabled(false);
        }

        // Setup event listeners
        setupEventListeners();

        // Setup Excel selection tracking
        trackExcelSelection();

        UIModule.setIdle();
    });
}

/**
 * Setup event listeners for buttons and tabs
 */
function setupEventListeners() {
    // Tab navigation
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            const tabName = e.target.getAttribute('data-tab');

            // Update active states
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            e.target.classList.add('active');
            document.getElementById(tabName)?.classList.add('active');
        });
    });

    // Process button
    const processButton = document.getElementById('processButton');
    if (processButton) {
        processButton.addEventListener('click', handleProcessSelection);
    }

    // Save settings button
    const saveButton = document.getElementById('saveButton');
    if (saveButton) {
        saveButton.addEventListener('click', handleSaveSettings);
    }

    // Clear API key button
    const clearButton = document.getElementById('clearButton');
    if (clearButton) {
        clearButton.addEventListener('click', handleClearApiKey);
    }
}

/**
 * Track Excel cell selection and update UI
 */
async function trackExcelSelection() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onSelectionChanged.add(onSelectionChanged);
            await context.sync();
        });
    } catch (error) {
        console.error('Error setting up selection tracking:', error);
    }

    // Initial update
    updateSelectedCellDisplay();
}

/**
 * Handle selection change event
 */
async function onSelectionChanged() {
    updateSelectedCellDisplay();
}

/**
 * Update the selected cell display
 */
async function updateSelectedCellDisplay() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedData(Excel.ValueType.values);
            const cellAddress = context.application.getSelection();
            
            // Get the currently selected range
            currentSelectedRange = context.workbook.getSelectedData(Excel.ValueType.values);
            
            // Update UI with cell address (simplified - shows last part)
            const activeCell = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            const selection = context.workbook.worksheets.getActiveWorksheet().getRange('A1').getOffsetRange(0, 0).getIntersection(context.application.getSelection());
            
            // Try to get the selection address
            try {
                const selectedRange = context.workbook.worksheets.getActiveWorksheet().getSelection();
                selectedRange.load('address');
                await context.sync();
                UIModule.updateSelectedCell(selectedRange.address);
            } catch (err) {
                // Fallback if address is not available
                console.log('Using fallback cell address');
                UIModule.updateSelectedCell('Selected');
            }
        });
    } catch (error) {
        console.error('Error updating selected cell:', error);
    }
}

/**
 * Handle Process Selection button click
 */
async function handleProcessSelection() {
    // Check if API key exists
    if (!StorageModule.hasApiKey()) {
        UIModule.setError('API Key required - visit Settings tab');
        return;
    }

    UIModule.setProcessing('Reading cell...');

    try {
        let cellValue = null;
        let targetRange = null;

        await Excel.run(async (context) => {
            // Get the selected range
            const range = context.workbook.worksheets.getActiveWorksheet().getSelection();
            range.load('values,address');
            await context.sync();

            cellValue = range.values[0][0];
            const cellAddress = range.address;

            // Get the offset range (one column to the right)
            targetRange = range.getOffsetRange(0, 1);

            if (!cellValue) {
                throw new Error('Selected cell is empty');
            }

            console.log(`Processing cell ${cellAddress} with value: ${cellValue}`);
        });

        // Validate cell value
        if (cellValue === null || cellValue === undefined || cellValue === '') {
            UIModule.setError('Selected cell is empty');
            return;
        }

        UIModule.setProcessing('Sending to AI...');

        // Get settings
        const apiKey = StorageModule.getApiKey();
        const model = StorageModule.getModel();
        const systemPrompt = StorageModule.getSystemPrompt();

        // Call API
        const response = await APIModule.callOpenRouter(
            cellValue.toString(),
            apiKey,
            model,
            systemPrompt
        );

        // Display response
        UIModule.displayResponse(response);
        UIModule.setProcessing('Writing to Excel...');

        // Write response to target cell
        await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getSelection();
            const targetRange = range.getOffsetRange(0, 1);
            
            targetRange.values = [[response]];
            targetRange.format.fill.color = '#E8F5E9';
            
            await context.sync();
        });

        UIModule.setSuccess('Response written to cell');
    } catch (error) {
        const errorMessage = APIModule.formatErrorMessage(error);
        UIModule.setError(errorMessage);
        console.error('Error processing selection:', error);
    }
}

/**
 * Handle Save Settings button click
 */
async function handleSaveSettings() {
    const settings = UIModule.getSettingsForm();

    // Validate input
    if (!settings.apiKey || settings.apiKey.trim() === '') {
        UIModule.showSettingsStatus('error', 'API Key cannot be empty');
        return;
    }

    // Save to storage
    const apiKeySaved = StorageModule.setApiKey(settings.apiKey);
    const modelSaved = StorageModule.setModel(settings.model);
    const promptSaved = StorageModule.setSystemPrompt(settings.systemPrompt);

    if (apiKeySaved && modelSaved && promptSaved) {
        UIModule.showSettingsStatus('success', 'Settings saved successfully');
        UIModule.setProcessButtonEnabled(true);
        UIModule.setIdle('API Key configured');
    } else {
        UIModule.showSettingsStatus('error', 'Failed to save settings');
    }
}

/**
 * Handle Clear API Key button click
 */
async function handleClearApiKey() {
    const confirmed = confirm('Are you sure you want to clear the API Key?');
    
    if (confirmed) {
        const cleared = StorageModule.clearApiKey();
        
        if (cleared) {
            UIModule.elements.apiKeyInput.value = '';
            UIModule.showSettingsStatus('success', 'API Key cleared');
            UIModule.setProcessButtonEnabled(false);
            UIModule.setError('API Key removed - add a new one to continue');
        } else {
            UIModule.showSettingsStatus('error', 'Failed to clear API Key');
        }
    }
}

/**
 * Document ready handler
 */
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});
