// Excel CRUD Manager - Main Application Logic with Enhanced Error Handling

Office.onReady((info) => {
    console.log('Office.onReady called', info);
    if (info.host === Office.HostType.Excel) {
        console.log('Excel host detected, initializing app...');
        setTimeout(initializeApp, 1000); // Add small delay for full initialization
    } else {
        console.log('Not running in Excel, host:', info.host);
        showStatus('This add-in requires Excel to function properly.', 'warning');
    }
});

// Global variables
let selectedRecord = null;
const WORKSHEET_NAME = 'CostItems';

// Initialize the application
function initializeApp() {
    try {
        console.log('Initializing app...');
        setupEventListeners();
        showStatus('Excel CRUD Manager loaded successfully!', 'success');
        
        // Setup worksheet with delay to ensure Excel is ready
        setTimeout(setupWorksheet, 2000);
        setTimeout(loadAllRecords, 3000);
    } catch (error) {
        console.error('Error in initializeApp:', error);
        showStatus('Error initializing app: ' + error.message, 'error');
    }
}

// Set up event listeners
function setupEventListeners() {
    try {
        console.log('Setting up event listeners...');
        
        const createForm = document.getElementById('createForm');
        const updateForm = document.getElementById('updateForm');
        
        if (createForm) {
            createForm.addEventListener('submit', handleCreateRecord);
            console.log('Create form listener added');
        }
        
        if (updateForm) {
            updateForm.addEventListener('submit', handleUpdateRecord);
            console.log('Update form listener added');
        }
        
    } catch (error) {
        console.error('Error setting up event listeners:', error);
        showStatus('Error setting up form handlers: ' + error.message, 'error');
    }
}

// Setup worksheet with headers if it doesn't exist
async function setupWorksheet() {
    try {
        console.log('Setting up worksheet...');
        
        await Excel.run(async (context) => {
            try {
                const worksheets = context.workbook.worksheets;
                worksheets.load('items/name');
                await context.sync();
                
                let worksheet = worksheets.items.find(sheet => sheet.name === WORKSHEET_NAME);
                
                if (!worksheet) {
                    console.log('Creating new worksheet:', WORKSHEET_NAME);
                    worksheet = worksheets.add(WORKSHEET_NAME);
                    
                    // Add headers
                    const headers = [
                        'ID', 'Item Name', 'Unit Cost', 'Item Type', 'Quantity',
                        'Total Cost', 'Approval Status', 'Requested By', 'Request Date',
                        'Category', 'Vendor', 'Description', 'Unit of Measurement',
                        'Is Active', 'Creation Date', 'Last Modified', 'Notes'
                    ];
                    
                    const headerRange = worksheet.getRange('A1:Q1');
                    headerRange.values = [headers];
                    headerRange.format.font.bold = true;
                    headerRange.format.fill.color = '#4F81BD';
                    headerRange.format.font.color = 'white';
                    
                    // Auto-fit columns
                    worksheet.getUsedRange().format.autofitColumns();
                    
                    await context.sync();
                    console.log('Worksheet created successfully');
                    showStatus('Worksheet "' + WORKSHEET_NAME + '" created successfully!', 'success');
                } else {
                    console.log('Worksheet already exists:', WORKSHEET_NAME);
                    showStatus('Connected to existing worksheet: ' + WORKSHEET_NAME, 'info');
                }
            } catch (innerError) {
                console.error('Error in Excel.run context:', innerError);
                throw innerError;
            }
        });
    } catch (error) {
        console.error('Error setting up worksheet:', error);
        showStatus('Note: Worksheet will be created when you add your first record.', 'info');
    }
}

// CREATE - Handle form submission for creating new records
async function handleCreateRecord(event) {
    event.preventDefault();
    console.log('Handling create record...');
    
    try {
        const formData = {
            itemName: document.getElementById('itemName').value.trim(),
            unitCost: parseFloat(document.getElementById('unitCost').value),
            itemType: document.getElementById('itemType').value,
            category: document.getElementById('category').value.trim(),
            vendor: document.getElementById('vendor').value.trim(),
            description: document.getElementById('description').value.trim()
        };
        
        console.log('Form data:', formData);
        
        // Validate form data
        if (!validateFormData(formData)) return;
        
        await Excel.run(async (context) => {
            try {
                // Ensure worksheet exists
                const worksheets = context.workbook.worksheets;
                worksheets.load('items/name');
                await context.sync();
                
                let worksheet = worksheets.items.find(sheet => sheet.name === WORKSHEET_NAME);
                if (!worksheet) {
                    worksheet = worksheets.add(WORKSHEET_NAME);
                    // Add headers for new worksheet
                    const headers = [
                        'ID', 'Item Name', 'Unit Cost', 'Item Type', 'Quantity',
                        'Total Cost', 'Approval Status', 'Requested By', 'Request Date',
                        'Category', 'Vendor', 'Description', 'Unit of Measurement',
                        'Is Active', 'Creation Date', 'Last Modified', 'Notes'
                    ];
                    const headerRange = worksheet.getRange('A1:Q1');
                    headerRange.values = [headers];
                    headerRange.format.font.bold = true;
                    headerRange.format.fill.color = '#4F81BD';
                    headerRange.format.font.color = 'white';
                    await context.sync();
                }
                
                const usedRange = worksheet.getUsedRange();
                usedRange.load('rowCount');
                await context.sync();
                
                const nextRow = usedRange.rowCount + 1;
                const itemId = generateItemId(nextRow - 1);
                const currentDate = new Date().toISOString();
                
                const newRowData = [
                    itemId,                                    // ID
                    formData.itemName,                         // Item Name
                    formData.unitCost,                         // Unit Cost
                    formData.itemType,                         // Item Type
                    1,                                         // Quantity (default)
                    formData.unitCost,                         // Total Cost
                    'Pending',                                 // Approval Status
                    'User',                                    // Requested By
                    currentDate,                               // Request Date
                    formData.category,                         // Category
                    formData.vendor,                           // Vendor
                    formData.description,                      // Description
                    'Each',                                    // Unit of Measurement
                    true,                                      // Is Active
                    currentDate,                               // Creation Date
                    currentDate,                               // Last Modified
                    ''                                         // Notes
                ];
                
                const newRowRange = worksheet.getRange(`A${nextRow}:Q${nextRow}`);
                newRowRange.values = [newRowData];
                
                // Format the new row
                newRowRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                newRowRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                
                await context.sync();
                
                console.log('Record created successfully:', itemId);
                showStatus(`Record created successfully! ID: ${itemId}`, 'success');
                clearForm('createForm');
                
                // Refresh data after a short delay
                setTimeout(loadAllRecords, 1000);
                
            } catch (innerError) {
                console.error('Error in create Excel.run context:', innerError);
                throw innerError;
            }
        });
    } catch (error) {
        console.error('Error creating record:', error);
        showStatus('Error creating record: ' + error.message, 'error');
    }
}

// READ - Load all records from the worksheet
async function loadAllRecords() {
    try {
        console.log('Loading all records...');
        
        await Excel.run(async (context) => {
            try {
                const worksheets = context.workbook.worksheets;
                worksheets.load('items/name');
                await context.sync();
                
                const worksheet = worksheets.items.find(sheet => sheet.name === WORKSHEET_NAME);
                if (!worksheet) {
                    console.log('Worksheet not found, showing empty state');
                    document.getElementById('recordsList').innerHTML = `
                        <div class="no-records">
                            <i class="fas fa-inbox"></i>
                            <p>No records found. Create your first record!</p>
                        </div>
                    `;
                    return;
                }
                
                const usedRange = worksheet.getUsedRange();
                if (!usedRange) {
                    console.log('No data in worksheet');
                    document.getElementById('recordsList').innerHTML = `
                        <div class="no-records">
                            <i class="fas fa-inbox"></i>
                            <p>No records found. Create your first record!</p>
                        </div>
                    `;
                    return;
                }
                
                usedRange.load('values');
                await context.sync();
                
                if (usedRange.values && usedRange.values.length > 1) {
                    console.log('Found', usedRange.values.length - 1, 'records');
                    displayRecords(usedRange.values);
                } else {
                    console.log('No data rows found');
                    document.getElementById('recordsList').innerHTML = `
                        <div class="no-records">
                            <i class="fas fa-inbox"></i>
                            <p>No records found. Create your first record!</p>
                        </div>
                    `;
                }
            } catch (innerError) {
                console.error('Error in load Excel.run context:', innerError);
                throw innerError;
            }
        });
    } catch (error) {
        console.error('Error loading records:', error);
        showStatus('Note: Records will appear here after you create some.', 'info');
        document.getElementById('recordsList').innerHTML = `
            <div class="no-records">
                <i class="fas fa-inbox"></i>
                <p>Create your first record to see it here!</p>
            </div>
        `;
    }
}

// Display records in the Read tab
function displayRecords(data) {
    try {
        console.log('Displaying', data.length - 1, 'records');
        const recordsList = document.getElementById('recordsList');
        if (!recordsList) {
            console.error('Records list element not found');
            return;
        }
        
        const headers = data[0];
        const records = data.slice(1);
        
        let html = '';
        
        records.forEach((record, index) => {
            if (record[0] && record[13]) { // ID exists and is_active = true
                const rowIndex = index + 2; // +2 because we sliced off header and arrays are 0-indexed
                html += `
                    <div class="record-card" onclick="selectRecord('${record[0]}', ${rowIndex}, [${record.map(r => `'${String(r).replace(/'/g, "\\'")}'`).join(',')}])">
                        <div class="record-header">
                            <h6>${record[1] || 'Unnamed Item'}</h6>
                            <span class="record-id">${record[0]}</span>
                        </div>
                        <div class="record-details">
                            <p><strong>Cost:</strong> ₦${(record[2] || 0).toLocaleString()}</p>
                            <p><strong>Type:</strong> ${record[3] || 'N/A'}</p>
                            <p><strong>Category:</strong> ${record[9] || 'N/A'}</p>
                            <p><strong>Status:</strong> 
                                <span class="status-badge status-${(record[6] || 'pending').toLowerCase()}">
                                    ${record[6] || 'Pending'}
                                </span>
                            </p>
                        </div>
                    </div>
                `;
            }
        });
        
        if (html === '') {
            html = `
                <div class="no-records">
                    <i class="fas fa-inbox"></i>
                    <p>No active records found.</p>
                </div>
            `;
        }
        
        recordsList.innerHTML = html;
        console.log('Records displayed successfully');
    } catch (error) {
        console.error('Error displaying records:', error);
        showStatus('Error displaying records: ' + error.message, 'error');
    }
}

// Utility function to validate form data
function validateFormData(data) {
    console.log('Validating form data:', data);
    
    if (!data.itemName) {
        showStatus('Item name is required', 'error');
        return false;
    }
    
    if (!data.unitCost || isNaN(data.unitCost) || data.unitCost <= 0) {
        showStatus('Valid unit cost is required', 'error');
        return false;
    }
    
    if (!data.itemType) {
        showStatus('Item type is required', 'error');
        return false;
    }
    
    console.log('Form data validation passed');
    return true;
}

// Generate unique item ID
function generateItemId(sequence) {
    const year = new Date().getFullYear();
    const paddedSequence = String(sequence).padStart(3, '0');
    return `ITEM${year}${paddedSequence}`;
}

// Clear form fields
function clearForm(formId) {
    try {
        const form = document.getElementById(formId);
        if (form) {
            form.reset();
            console.log('Form cleared:', formId);
        }
    } catch (error) {
        console.error('Error clearing form:', error);
    }
}

// Show status messages
function showStatus(message, type) {
    try {
        console.log('Status:', type, '-', message);
        
        const container = document.getElementById('statusContainer');
        if (!container) {
            console.error('Status container not found');
            return;
        }
        
        const messageId = 'status-' + Date.now();
        
        const messageDiv = document.createElement('div');
        messageDiv.id = messageId;
        messageDiv.className = `status-message status-${type}`;
        messageDiv.innerHTML = `
            <div class="d-flex justify-content-between align-items-center">
                <span>${message}</span>
                <button type="button" class="btn-close" onclick="document.getElementById('${messageId}').remove()" aria-label="Close">×</button>
            </div>
        `;
        
        container.appendChild(messageDiv);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            const element = document.getElementById(messageId);
            if (element) {
                element.remove();
            }
        }, 5000);
    } catch (error) {
        console.error('Error showing status:', error);
    }
}

// Handle record selection for update/delete
function selectRecord(recordId, rowIndex, recordData) {
    try {
        console.log('Record selected:', recordId);
        
        // Remove previous selection
        document.querySelectorAll('.record-card').forEach(card => {
            card.classList.remove('selected');
        });
        
        // Add selection to clicked card
        if (event && event.currentTarget) {
            event.currentTarget.classList.add('selected');
        }
        
        // Store selected record data
        selectedRecord = {
            id: recordId,
            row: rowIndex,
            data: recordData
        };
        
        showStatus(`Record ${recordId} selected for editing`, 'info');
    } catch (error) {
        console.error('Error selecting record:', error);
        showStatus('Error selecting record: ' + error.message, 'error');
    }
}

// Update and Delete functions would go here...
// For now, let's focus on getting Create and Read working

// Filter records based on search input
function filterRecords() {
    try {
        const searchTerm = document.getElementById('searchInput').value.toLowerCase();
        const recordCards = document.querySelectorAll('.record-card');
        
        recordCards.forEach(card => {
            const text = card.textContent.toLowerCase();
            card.style.display = text.includes(searchTerm) ? 'block' : 'none';
        });
    } catch (error) {
        console.error('Error filtering records:', error);
    }
}

// Log function for debugging
function log(message) {
    console.log('[Excel CRUD Manager]', message);
}

// Global error handler
window.addEventListener('error', function(event) {
    console.error('Global error caught:', event.error);
    showStatus('An unexpected error occurred. Check console for details.', 'error');
});

// Add this to help debug
console.log('App.js loaded successfully');
