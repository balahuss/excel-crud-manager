// Excel CRUD Manager - Main Application Logic

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Office Add-in loaded successfully');
        initializeApp();
    }
});

// Global variables
let selectedRecord = null;
const WORKSHEET_NAME = 'CostItems';

// Initialize the application
function initializeApp() {
    setupEventListeners();
    setupWorksheet();
    loadAllRecords();
    showStatus('Excel CRUD Manager loaded successfully!', 'success');
}

// Set up event listeners
function setupEventListeners() {
    document.getElementById('createForm').addEventListener('submit', handleCreateRecord);
    document.getElementById('updateForm').addEventListener('submit', handleUpdateRecord);
}

// Setup worksheet with headers if it doesn't exist
async function setupWorksheet() {
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load('items/name');
            await context.sync();
            
            let worksheet = worksheets.items.find(sheet => sheet.name === WORKSHEET_NAME);
            
            if (!worksheet) {
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
                console.log('Worksheet created with headers');
            }
        });
    } catch (error) {
        console.error('Error setting up worksheet:', error);
        showStatus('Error setting up worksheet: ' + error.message, 'error');
    }
}

// CREATE - Handle form submission for creating new records
async function handleCreateRecord(event) {
    event.preventDefault();
    
    const formData = {
        itemName: document.getElementById('itemName').value.trim(),
        unitCost: parseFloat(document.getElementById('unitCost').value),
        itemType: document.getElementById('itemType').value,
        category: document.getElementById('category').value.trim(),
        vendor: document.getElementById('vendor').value.trim(),
        description: document.getElementById('description').value.trim()
    };
    
    // Validate form data
    if (!validateFormData(formData)) return;
    
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(WORKSHEET_NAME);
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
            
            showStatus(`Record created successfully! ID: ${itemId}`, 'success');
            clearForm('createForm');
            loadAllRecords();
        });
    } catch (error) {
        console.error('Error creating record:', error);
        showStatus('Error creating record: ' + error.message, 'error');
    }
}

// READ - Load all records from the worksheet
async function loadAllRecords() {
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(WORKSHEET_NAME);
            const usedRange = worksheet.getUsedRange();
            usedRange.load('values');
            await context.sync();
            
            if (usedRange.values && usedRange.values.length > 1) {
                displayRecords(usedRange.values);
            } else {
                document.getElementById('recordsList').innerHTML = `
                    <div class="no-records">
                        <i class="fas fa-inbox"></i>
                        <p>No records found. Create your first record!</p>
                    </div>
                `;
            }
        });
    } catch (error) {
        console.error('Error loading records:', error);
        showStatus('Error loading records: ' + error.message, 'error');
    }
}

// Display records in the Read tab
function displayRecords(data) {
    const recordsList = document.getElementById('recordsList');
    const headers = data[0];
    const records = data.slice(1);
    
    let html = '';
    
    records.forEach((record, index) => {
        if (record[0] && record[13]) { // ID exists and is_active = true
            const rowIndex = index + 2; // +2 because we sliced off header and arrays are 0-indexed
            html += `
                <div class="record-card" onclick="selectRecord('${record[0]}', ${rowIndex}, [${record.map(r => `'${r}'`).join(',')}])">
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
}

// Handle record selection for update/delete
function selectRecord(recordId, rowIndex, recordData) {
    // Remove previous selection
    document.querySelectorAll('.record-card').forEach(card => {
        card.classList.remove('selected');
    });
    
    // Add selection to clicked card
    event.currentTarget.classList.add('selected');
    
    // Store selected record data
    selectedRecord = {
        id: recordId,
        row: rowIndex,
        data: recordData
    };
    
    // Populate update form
    populateUpdateForm(recordData);
    populateDeletePreview(recordData);
    
    showStatus(`Record ${recordId} selected for editing`, 'info');
}

// Populate the update form with selected record data
function populateUpdateForm(data) {
    document.getElementById('updateRecordId').value = data[0];
    document.getElementById('updateRowIndex').value = selectedRecord.row;
    document.getElementById('updateItemName').value = data[1] || '';
    document.getElementById('updateUnitCost').value = data[2] || '';
    document.getElementById('updateItemType').value = data[3] || '';
    document.getElementById('updateCategory').value = data[9] || '';
    document.getElementById('updateVendor').value = data[10] || '';
    document.getElementById('updateDescription').value = data[11] || '';
    
    document.getElementById('updateInstructions').style.display = 'none';
    document.getElementById('updateForm').style.display = 'block';
}

// UPDATE - Handle form submission for updating records
async function handleUpdateRecord(event) {
    event.preventDefault();
    
    if (!selectedRecord) {
        showStatus('Please select a record to update', 'warning');
        return;
    }
    
    const formData = {
        itemName: document.getElementById('updateItemName').value.trim(),
        unitCost: parseFloat(document.getElementById('updateUnitCost').value),
        itemType: document.getElementById('updateItemType').value,
        category: document.getElementById('updateCategory').value.trim(),
        vendor: document.getElementById('updateVendor').value.trim(),
        description: document.getElementById('updateDescription').value.trim()
    };
    
    if (!validateFormData(formData)) return;
    
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(WORKSHEET_NAME);
            const updateRange = worksheet.getRange(`B${selectedRecord.row}:Q${selectedRecord.row}`);
            
            const updatedData = [
                formData.itemName,                         // Item Name
                formData.unitCost,                         // Unit Cost
                formData.itemType,                         // Item Type
                selectedRecord.data[4],                    // Keep existing quantity
                formData.unitCost * (selectedRecord.data[4] || 1), // Recalculate total
                selectedRecord.data[6],                    // Keep existing approval status
                selectedRecord.data[7],                    // Keep existing requested by
                selectedRecord.data[8],                    // Keep existing request date
                formData.category,                         // Category
                formData.vendor,                           // Vendor
                formData.description,                      // Description
                selectedRecord.data[12],                   // Keep existing unit of measurement
                true,                                      // Keep active
                selectedRecord.data[14],                   // Keep existing creation date
                new Date().toISOString(),                  // Update last modified
                selectedRecord.data[16]                    // Keep existing notes
            ];
            
            updateRange.values = [updatedData];
            await context.sync();
            
            showStatus(`Record ${selectedRecord.id} updated successfully!`, 'success');
            clearSelection();
            loadAllRecords();
        });
    } catch (error) {
        console.error('Error updating record:', error);
        showStatus('Error updating record: ' + error.message, 'error');
    }
}

// Populate the delete preview
function populateDeletePreview(data) {
    const preview = document.getElementById('deletePreview');
    preview.innerHTML = `
        <div class="delete-preview-card">
            <h5><i class="fas fa-exclamation-triangle text-danger"></i> Confirm Deletion</h5>
            <div class="record-summary">
                <p><strong>ID:</strong> ${data[0]}</p>
                <p><strong>Name:</strong> ${data[1]}</p>
                <p><strong>Cost:</strong> ₦${(data[2] || 0).toLocaleString()}</p>
                <p><strong>Type:</strong> ${data[3]}</p>
                <p><strong>Category:</strong> ${data[9] || 'N/A'}</p>
            </div>
            <div class="d-flex gap-2">
                <button onclick="confirmDelete()" class="btn btn-danger">
                    <i class="fas fa-trash"></i> Confirm Delete
                </button>
                <button onclick="clearSelection()" class="btn btn-secondary">
                    <i class="fas fa-times"></i> Cancel
                </button>
            </div>
        </div>
    `;
    
    document.getElementById('deleteInstructions').style.display = 'none';
    preview.style.display = 'block';
}

// DELETE - Confirm and execute deletion (soft delete)
async function confirmDelete() {
    if (!selectedRecord) {
        showStatus('No record selected for deletion', 'warning');
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(WORKSHEET_NAME);
            const activeCell = worksheet.getRange(`N${selectedRecord.row}`);
            activeCell.values = [[false]]; // Set is_active to false
            
            await context.sync();
            
            showStatus(`Record ${selectedRecord.id} deleted successfully!`, 'success');
            clearSelection();
            loadAllRecords();
        });
    } catch (error) {
        console.error('Error deleting record:', error);
        showStatus('Error deleting record: ' + error.message, 'error');
    }
}

// Clear selection and reset forms
function clearSelection() {
    selectedRecord = null;
    
    // Clear visual selection
    document.querySelectorAll('.record-card').forEach(card => {
        card.classList.remove('selected');
    });
    
    // Reset update form
    document.getElementById('updateInstructions').style.display = 'block';
    document.getElementById('updateForm').style.display = 'none';
    document.getElementById('updateForm').reset();
    
    // Reset delete preview
    document.getElementById('deleteInstructions').style.display = 'block';
    document.getElementById('deletePreview').style.display = 'none';
    
    showStatus('Selection cleared', 'info');
}

// Utility function to validate form data
function validateFormData(data) {
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
    document.getElementById(formId).reset();
}

// Show status messages
function showStatus(message, type) {
    const container = document.getElementById('statusContainer');
    const messageId = 'status-' + Date.now();
    
    const messageDiv = document.createElement('div');
    messageDiv.id = messageId;
    messageDiv.className = `status-message status-${type}`;
    messageDiv.innerHTML = `
        <div class="d-flex justify-content-between align-items-center">
            <span>${message}</span>
            <button type="button" class="btn-close btn-close-white" onclick="document.getElementById('${messageId}').remove()"></button>
        </div>
    `;
    
    container.appendChild(messageDiv);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (document.getElementById(messageId)) {
            document.getElementById(messageId).remove();
        }
    }, 5000);
}

// Filter records based on search input
function filterRecords() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const recordCards = document.querySelectorAll('.record-card');
    
    recordCards.forEach(card => {
        const text = card.textContent.toLowerCase();
        card.style.display = text.includes(searchTerm) ? 'block' : 'none';
    });
}

// Log function for debugging
function log(message) {
    console.log('[Excel CRUD Manager]', message);
}
