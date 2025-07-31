// Initialize the application when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
    // Initialize data storage
    initializeDataStorage();
    
    // Set up event listeners
    setupEventListeners();
    
    // Add the first row by default
    addNewRow();
    
    // Update the preview
    updatePreview();
    
    // Initialize Excel editor
    initializeExcelEditor();
});

// Data storage for suppliers and stock codes
let suppliers = [];
let stockCodes = [];

// Function to initialize data storage from localStorage
function initializeDataStorage() {
    // Load suppliers from localStorage or set defaults
    const savedSuppliers = localStorage.getItem('qrLabelSuppliers');
    if (savedSuppliers) {
        suppliers = JSON.parse(savedSuppliers);
    } else {
        // Default suppliers
        suppliers = ['Supplier A', 'Supplier B', 'Supplier C'];
        localStorage.setItem('qrLabelSuppliers', JSON.stringify(suppliers));
    }
    
    // Load stock codes from localStorage or set defaults
    const savedStockCodes = localStorage.getItem('qrLabelStockCodes');
    if (savedStockCodes) {
        stockCodes = JSON.parse(savedStockCodes);
    } else {
        // Default stock codes
        stockCodes = ['SC001', 'SC002', 'SC003'];
        localStorage.setItem('qrLabelStockCodes', JSON.stringify(stockCodes));
    }
    
    // Populate the data management lists
    updateSupplierList();
    updateStockCodeList();
}

// Function to set up all event listeners
function setupEventListeners() {
    // Add row button
    document.getElementById('addRowBtn').addEventListener('click', addNewRow);
    
    // Generate PDF button
    document.getElementById('generatePdfBtn').addEventListener('click', generatePDF);
    
    // Add supplier button
    document.getElementById('addSupplierBtn').addEventListener('click', function() {
        const input = document.getElementById('newSupplierInput');
        if (input.value.trim()) {
            addSupplier(input.value.trim());
            input.value = '';
        }
    });
    
    // Add stock code button
    document.getElementById('addStockCodeBtn').addEventListener('click', function() {
        const input = document.getElementById('newStockCodeInput');
        if (input.value.trim()) {
            addStockCode(input.value.trim());
            input.value = '';
        }
    });
    
    // Note: Tab buttons use onclick in HTML, so we don't need event listeners here
    
    // Excel editor buttons
    document.getElementById('loadExcelBtn').addEventListener('click', loadExcelData);
    document.getElementById('saveExcelBtn').addEventListener('click', saveExcelData);
    document.getElementById('addExcelRowBtn').addEventListener('click', addExcelRow);
}

// Function to open a tab
function openTab(tabName) {
    // Hide all tab contents
    const tabContents = document.getElementsByClassName('tab-content');
    for (let i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove('active');
    }
    
    // Remove active class from all tab buttons
    const tabButtons = document.getElementsByClassName('tab-btn');
    for (let i = 0; i < tabButtons.length; i++) {
        tabButtons[i].classList.remove('active');
    }
    
    // Show the selected tab content
    document.getElementById(tabName).classList.add('active');
    
    // Mark the button that was clicked as active (for onclick from HTML)
    if (event && event.currentTarget) {
        event.currentTarget.classList.add('active');
    }
}

// Function to add a new row to the label table
function addNewRow() {
    const tbody = document.getElementById('labelTableBody');
    const rowCount = tbody.children.length;
    
    // Create a new row
    const row = document.createElement('tr');
    
    // Get today's date in YYYY-MM-DD format
    const today = new Date();
    const formattedDate = today.toISOString().split('T')[0];
    
    // Default values
    const defaultSupplier = suppliers.length > 0 ? suppliers[0] : '';
    const defaultStockCode = stockCodes.length > 0 ? stockCodes[0] : '';
    
    // If there are existing rows, copy values from the first row
    let supplierValue = defaultSupplier;
    let stockCodeValue = defaultStockCode;
    let dateValue = formattedDate;
    
    if (rowCount > 0) {
        const firstRow = tbody.children[0];
        const supplierSelect = firstRow.querySelector('.supplier-select');
        const stockCodeSelect = firstRow.querySelector('.stock-code-select');
        const dateInput = firstRow.querySelector('.date-input');
        
        if (supplierSelect) supplierValue = supplierSelect.value;
        if (stockCodeSelect) stockCodeValue = stockCodeSelect.value;
        if (dateInput) dateValue = dateInput.value;
    }
    
    // Create the row HTML
    row.innerHTML = `
        <td>
            <select class="supplier-select" onchange="updatePreview()">
                ${suppliers.map(supplier => 
                    `<option value="${supplier}" ${supplier === supplierValue ? 'selected' : ''}>${supplier}</option>`
                ).join('')}
            </select>
        </td>
        <td>
            <select class="stock-code-select" onchange="updatePreview()">
                ${stockCodes.map(code => 
                    `<option value="${code}" ${code === stockCodeValue ? 'selected' : ''}>${code}</option>`
                ).join('')}
            </select>
        </td>
        <td>
            <input type="number" class="mass-input" placeholder="Enter mass" step="0.01" min="0" onchange="updatePreview()">
        </td>
        <td>
            <input type="date" class="date-input" value="${dateValue}" onchange="updatePreview()">
        </td>
        <td>
            <button class="action-btn delete-btn" onclick="deleteRow(this)">Delete</button>
        </td>
    `;
    
    // Add the row to the table
    tbody.appendChild(row);
    
    // Update the preview
    updatePreview();
}

// Function to delete a row
function deleteRow(button) {
    const row = button.closest('tr');
    row.remove();
    updatePreview();
}

// Function to update data (previously updatePreview)
function updatePreview() {
    // This function is kept for compatibility but no longer updates previews
    // It's still called when form values change
    // No action needed as preview functionality has been removed
}

// Function to get data from a table row
function getRowData(row) {
    const supplierSelect = row.querySelector('.supplier-select');
    const stockCodeSelect = row.querySelector('.stock-code-select');
    const massInput = row.querySelector('.mass-input');
    const dateInput = row.querySelector('.date-input');
    
    return {
        supplier: supplierSelect.value,
        stockCode: stockCodeSelect.value,
        mass: massInput.value || '0',
        date: dateInput.value
    };
}

// Function to create a label element from row data
function createLabelElement(rowData, existingUuid = null) {
    // Create a new label element
    const label = document.createElement('div');
    label.className = 'label';
    
    // Clone the label content template
    const template = document.getElementById('labelTemplate');
    const content = template.querySelector('.label-content').cloneNode(true);
    
    // Set the label information
    content.querySelector('.supplier-value').textContent = rowData.supplier;
    content.querySelector('.stock-code-value').textContent = rowData.stockCode;
    content.querySelector('.mass-value').textContent = rowData.mass;
    content.querySelector('.date-value').textContent = formatDate(rowData.date);
    // Type is already set to 'Raw' in the HTML template
    
    // Use existing UUID if provided, otherwise generate a new one
    const labelUuid = existingUuid || uuid.v4();
    
    // Generate QR code with data in JSON format for better structure
    // Create a JSON object with the label data including UUID (only in QR code, not displayed)
    const qrCodeData = JSON.stringify({
        supplier: rowData.supplier,
        stockCode: rowData.stockCode,
        mass: rowData.mass,
        date: formatDate(rowData.date),
        type: "Raw", // Add type field as requested
        uuid: labelUuid // Add UUID to QR code data but not to visible label
    });
    
    // Store the UUID in a data attribute on the label element for future reference
    label.setAttribute('data-uuid', labelUuid);
    const qrCodeElement = content.querySelector('.qr-code');
    
    try {
        // Create QR code using the library
        const typeNumber = 0; // Auto-detect QR code type number (0 = auto)
        const errorCorrectionLevel = 'M'; // Medium error correction level for JSON data
        
        // Create QR code directly using the global qrcode object
        const qr = qrcode(typeNumber, errorCorrectionLevel);
        qr.addData(qrCodeData);
        qr.make();
        
        // Create an image element for the QR code
        const qrImg = document.createElement('img');
        // Use a cell size of 8 and margin of 1 for JSON data
        qrImg.src = qr.createImgTag(8, 1).match(/src="([^"]+)"/)[1]; // Extract src from img tag
        // Set explicit width and height to ensure proper display
        qrImg.style.width = '100%';
        qrImg.style.height = '100%';
        qrImg.style.objectFit = 'contain';
        
        // Clear any previous content
        qrCodeElement.innerHTML = '';
        qrCodeElement.appendChild(qrImg);
        
        // Add data attributes with the QR code data for debugging
        qrImg.setAttribute('data-qr-content', qrCodeData);
        qrImg.setAttribute('data-qr-format', 'json');
        console.log('QR Code Data (JSON):', qrCodeData); // Log data for debugging
    } catch (error) {
        console.error('Error generating QR code:', error);
        qrCodeElement.textContent = 'QR Code Error: ' + error.message;
    }
    
    // Add the content to the label
    label.appendChild(content);
    
    return label;
}

// Function to format date for display in dd/mm/yyyy format
function formatDate(dateString) {
    if (!dateString) return '';
    
    const date = new Date(dateString);
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
}

// Function to generate a PDF with all labels and export CSV
function generatePDF() {
    const { jsPDF } = window.jspdf;
    
    // Get all rows from the table
    const rows = document.querySelectorAll('#labelTableBody tr');
    
    if (rows.length === 0) {
        alert('Please add at least one row before generating a PDF.');
        return;
    }
    
    // Get first row data for filename
    const firstRowData = getRowData(rows[0]);
    const formattedDate = formatDate(firstRowData.date);
    const supplierName = firstRowData.supplier.replace(/[/\:*?"<>|]/g, '_'); // Remove invalid filename characters
    
    // Generate CSV data with UUIDs for tracking
    generateCSV(rows);
    
    // Create a temporary container for the labels
    const tempContainer = document.createElement('div');
    tempContainer.style.position = 'absolute';
    tempContainer.style.left = '-9999px';
    tempContainer.style.top = '-9999px';
    document.body.appendChild(tempContainer);
    
    // Add all labels to the temporary container
    rows.forEach(row => {
        const rowData = getRowData(row);
        // Get UUID from the row if it exists
        const existingUuid = row.getAttribute('data-uuid');
        // Create label with the existing UUID or generate a new one
        const label = createLabelElement(rowData, existingUuid);
        tempContainer.appendChild(label);
    });
    
    // Create a new PDF document optimized for Zebra ZT230 label printer (100mm x 75mm)
    const pdf = new jsPDF({
        orientation: 'landscape', // Landscape for 100x75mm labels
        unit: 'mm',
        format: [100, 75], // Exact dimensions for Zebra ZT230 labels
        compress: true // Compress the PDF for better quality
    });
    
    // Define label dimensions (100mm x 75mm)
    const labelWidth = 100;
    const labelHeight = 75;
    
    // Process each label
    const labels = tempContainer.querySelectorAll('.label');
    
    const processLabels = async () => {
        // Process each label on its own page for the Zebra printer
        for (let i = 0; i < labels.length; i++) {
            const label = labels[i];
            const labelUuid = label.getAttribute('data-uuid');
            
            // Add a new page for each label after the first one
            if (i > 0) {
                pdf.addPage([100, 75]); // Add a new page with the correct label size
            }
            
            // Log the UUID for tracking purposes
            console.log(`Processing label with UUID: ${labelUuid}`);
            
            // Convert the label to an image with higher quality settings
            const canvas = await html2canvas(label, {
                scale: 5, // Higher scale for better print quality
                useCORS: true,
                logging: false,
                backgroundColor: '#FFFFFF', // Ensure white background
                imageTimeout: 0, // No timeout for image loading
                allowTaint: true // Allow cross-origin images
            });
            
            // Add the image to the PDF - fill the entire label area with a small margin
            const imgData = canvas.toDataURL('image/jpeg', 1.0);
            // Add a small margin (1mm) to prevent content from being cut off at edges
            pdf.addImage(imgData, 'JPEG', 1, 1, labelWidth - 2, labelHeight - 2);
        }
        
        // Create the Generated Labels folder if it doesn't exist
        ensureFolderExists('Generated Labels');
        
        // Save the PDF with supplier name and date
        // PDF save
        const pdfFileName = `${supplierName}_${formattedDate}.pdf`;
        // Save the PDF to the Generated Labels folder on the server (if running locally)
        saveFileToProjectFolder('Generated Labels/' + pdfFileName, pdf.output('blob'));
        pdf.save(pdfFileName);
        
        // Excel save
        // Save the Excel file to the Receiving schedule folder on the server (if running locally)
    saveFileToProjectFolder('Receiving schedule/Receiving Schedule.xlsx', blob);
    link.setAttribute('download', 'Receiving Schedule.xlsx');
        
        // Remove the temporary container
        document.body.removeChild(tempContainer);
    };
    
    processLabels();
}

// Function to save a file to the project folder (requires server-side support)
function saveFileToProjectFolder(path, blob) {
    // This function requires server-side support to actually save files to disk
    // For now, it is a placeholder to show intent
    // In a real Electron app or with a backend, you would POST the blob to the server here
    console.log('Saving file to project folder:', path);
}

// Function to ensure a folder exists
function ensureFolderExists(folderName) {
    // This is a client-side function that doesn't actually create folders on the file system
    // It's just a placeholder for the concept, as browsers can't directly create folders
    // The actual folder creation happens when the user selects the download location
    console.log(`Ensuring folder exists: ${folderName}`);
    // In a real application with server-side code, you would create the folder here
}

// Function to add a new supplier
function addSupplier(name) {
    if (suppliers.includes(name)) {
        alert('This supplier already exists.');
        return;
    }
    
    suppliers.push(name);
    localStorage.setItem('qrLabelSuppliers', JSON.stringify(suppliers));
    updateSupplierList();
    updateSupplierDropdowns();
}

// Function to add a new stock code
function addStockCode(code) {
    if (stockCodes.includes(code)) {
        alert('This stock code already exists.');
        return;
    }
    
    stockCodes.push(code);
    localStorage.setItem('qrLabelStockCodes', JSON.stringify(stockCodes));
    updateStockCodeList();
    updateStockCodeDropdowns();
}

// Function to remove a supplier
function removeSupplier(index) {
    suppliers.splice(index, 1);
    localStorage.setItem('qrLabelSuppliers', JSON.stringify(suppliers));
    updateSupplierList();
    updateSupplierDropdowns();
}

// Function to remove a stock code
function removeStockCode(index) {
    stockCodes.splice(index, 1);
    localStorage.setItem('qrLabelStockCodes', JSON.stringify(stockCodes));
    updateStockCodeList();
    updateStockCodeDropdowns();
}

// Function to update the supplier list in the data management tab
function updateSupplierList() {
    const list = document.getElementById('supplierList');
    list.innerHTML = '';
    
    suppliers.forEach((supplier, index) => {
        const li = document.createElement('li');
        li.innerHTML = `
            ${supplier}
            <span class="delete-item" onclick="removeSupplier(${index})">&times;</span>
        `;
        list.appendChild(li);
    });
}

// Function to update the stock code list in the data management tab
function updateStockCodeList() {
    const list = document.getElementById('stockCodeList');
    list.innerHTML = '';
    
    stockCodes.forEach((code, index) => {
        const li = document.createElement('li');
        li.innerHTML = `
            ${code}
            <span class="delete-item" onclick="removeStockCode(${index})">&times;</span>
        `;
        list.appendChild(li);
    });
}

// Function to update all supplier dropdowns in the table
function updateSupplierDropdowns() {
    const selects = document.querySelectorAll('.supplier-select');
    
    selects.forEach(select => {
        const currentValue = select.value;
        select.innerHTML = '';
        
        suppliers.forEach(supplier => {
            const option = document.createElement('option');
            option.value = supplier;
            option.textContent = supplier;
            if (supplier === currentValue) {
                option.selected = true;
            }
            select.appendChild(option);
        });
        
        // If the current value is no longer in the list, select the first option
        if (!suppliers.includes(currentValue) && suppliers.length > 0) {
            select.value = suppliers[0];
        }
    });
    
    updatePreview();
}

// Function to update all stock code dropdowns in the table
function updateStockCodeDropdowns() {
    const selects = document.querySelectorAll('.stock-code-select');
    
    selects.forEach(select => {
        const currentValue = select.value;
        select.innerHTML = '';
        
        stockCodes.forEach(code => {
            const option = document.createElement('option');
            option.value = code;
            option.textContent = code;
            if (code === currentValue) {
                option.selected = true;
            }
            select.appendChild(option);
        });
        
        // If the current value is no longer in the list, select the first option
        if (!stockCodes.includes(currentValue) && stockCodes.length > 0) {
            select.value = stockCodes[0];
        }
    });
    
    updatePreview();
}

// Function to generate and append data to Excel workbook
function generateCSV(rows, appendOnly = true) {
    // Ensure the Receiving schedule folder exists
    ensureFolderExists('Receiving schedule');
    
    // Get existing workbook from localStorage or create a new one
    let workbook;
    const existingWorkbook = localStorage.getItem('receivingScheduleWorkbook');
    
    if (existingWorkbook && appendOnly) {
        // Parse existing workbook from base64 string
        try {
            const binaryString = atob(existingWorkbook);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            workbook = XLSX.read(bytes.buffer, { type: 'array' });
        } catch (e) {
            console.error('Error parsing existing workbook:', e);
            workbook = XLSX.utils.book_new();
        }
    } else {
        // Create a new workbook
        workbook = XLSX.utils.book_new();
    }
    
    // Get or create the worksheet
    let worksheet;
    if (workbook.SheetNames.includes('Receiving Schedule') && appendOnly) {
        worksheet = workbook.Sheets['Receiving Schedule'];
        
        // Check if the worksheet already has a UUID column, if not, update the headers
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:E1');
        if (range.e.c < 5) { // If there are less than 6 columns (0-indexed)
            // Add UUID header to existing worksheet
            XLSX.utils.sheet_add_aoa(worksheet, [['UUID']], { origin: { r: 0, c: 5 } });
        }
    } else {
        // Create a new worksheet with headers including UUID
        worksheet = XLSX.utils.aoa_to_sheet([['Supplier', 'Stock Code', 'Mass (kg)', 'Date', 'Timestamp', 'UUID']]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Receiving Schedule');
    }
    
    if (appendOnly) {
        // Get the current data range
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:F1'); // Updated to include UUID column
        let startRow = range.e.r + 1; // Start after the last row
        
        // Add new data from each row
        rows.forEach((row, index) => {
            const rowData = getRowData(row);
            const formattedDate = formatDate(rowData.date);
            const timestamp = new Date().toLocaleString();
            
            // Generate a UUID for this row if not already created
            // This will be used when creating the label and QR code
            const labelUuid = uuid.v4();
            
            // Add data to the worksheet including UUID
            XLSX.utils.sheet_add_aoa(worksheet, [[
                rowData.supplier,
                rowData.stockCode,
                rowData.mass,
                formattedDate,
                timestamp,
                labelUuid
            ]], { origin: startRow + index });
            
            // Store the UUID in a data attribute on the row for reference
            row.setAttribute('data-uuid', labelUuid);
        });
    }
    
    // Update the workbook in localStorage
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const base64 = btoa(String.fromCharCode.apply(null, new Uint8Array(wbout)));
    localStorage.setItem('receivingScheduleWorkbook', base64);
    
    // Create a Blob with the Excel data for download
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    
    // Create a download link
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    // Set link properties
    link.setAttribute('href', url);
    link.setAttribute('download', 'Receiving Schedule.xlsx');
    link.style.visibility = 'hidden';
    
    // Add to document, trigger download, and clean up
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    return blob; // Return the blob for further processing if needed
}

// No helper function needed as we're using the library's global qrcode function directly

// Excel Editor Functions

// Function to initialize the Excel editor
function initializeExcelEditor() {
    // This will be called when the page loads
    // We don't automatically load the data to avoid performance issues
    // User can click the Load button when they want to edit
}

// Function to load Excel data from localStorage
function loadExcelData() {
    const existingWorkbook = localStorage.getItem('receivingScheduleWorkbook');
    
    if (!existingWorkbook) {
        alert('No schedule data found. Create some labels first.');
        return;
    }
    
    try {
        // Parse existing workbook from base64 string
        const binaryString = atob(existingWorkbook);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        const workbook = XLSX.read(bytes.buffer, { type: 'array' });
        
        // Get the worksheet
        if (!workbook.SheetNames.includes('Receiving Schedule')) {
            alert('No schedule data found in the workbook.');
            return;
        }
        
        const worksheet = workbook.Sheets['Receiving Schedule'];
        
        // Convert worksheet to JSON
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        // Display data in the table
        displayExcelData(data);
        
    } catch (e) {
        console.error('Error parsing workbook:', e);
        alert('Error loading schedule data. Please try again.');
    }
}

// Function to display Excel data in the table
function displayExcelData(data) {
    const tbody = document.getElementById('excelTableBody');
    tbody.innerHTML = '';
    
    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        
        // Create cells for each column
        tr.innerHTML = `
            <td>
                <select class="excel-supplier-select">
                    ${suppliers.map(supplier => 
                        `<option value="${supplier}" ${supplier === row.Supplier ? 'selected' : ''}>${supplier}</option>`
                    ).join('')}
                </select>
            </td>
            <td>
                <select class="excel-stock-code-select">
                    ${stockCodes.map(code => 
                        `<option value="${code}" ${code === row['Stock Code'] ? 'selected' : ''}>${code}</option>`
                    ).join('')}
                </select>
            </td>
            <td>
                <input type="number" class="excel-mass-input" value="${row['Mass (kg)'] || ''}" step="0.01" min="0">
            </td>
            <td>
                <input type="date" class="excel-date-input" value="${formatDateForInput(row.Date)}">
            </td>
            <td>${row.Timestamp || ''}</td>
            <td>
                <button class="action-btn delete-btn" onclick="deleteExcelRow(this)">Delete</button>
            </td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Switch to the Excel tab
    openTab('editExcel');
}

// Function to format date from dd/mm/yyyy to yyyy-mm-dd for input fields
function formatDateForInput(dateString) {
    if (!dateString) return '';
    
    // Check if the date is in dd/mm/yyyy format
    const parts = dateString.split('/');
    if (parts.length === 3) {
        const day = parts[0];
        const month = parts[1];
        const year = parts[2];
        return `${year}-${month}-${day}`;
    }
    
    return dateString; // Return as is if not in expected format
}

// Function to add a new row to the Excel table
function addExcelRow() {
    const tbody = document.getElementById('excelTableBody');
    const tr = document.createElement('tr');
    
    // Get today's date in YYYY-MM-DD format
    const today = new Date();
    const formattedDate = today.toISOString().split('T')[0];
    
    // Default values
    const defaultSupplier = suppliers.length > 0 ? suppliers[0] : '';
    const defaultStockCode = stockCodes.length > 0 ? stockCodes[0] : '';
    
    // Create cells for each column
    tr.innerHTML = `
        <td>
            <select class="excel-supplier-select">
                ${suppliers.map(supplier => 
                    `<option value="${supplier}" ${supplier === defaultSupplier ? 'selected' : ''}>${supplier}</option>`
                ).join('')}
            </select>
        </td>
        <td>
            <select class="excel-stock-code-select">
                ${stockCodes.map(code => 
                    `<option value="${code}" ${code === defaultStockCode ? 'selected' : ''}>${code}</option>`
                ).join('')}
            </select>
        </td>
        <td>
            <input type="number" class="excel-mass-input" value="" step="0.01" min="0">
        </td>
        <td>
            <input type="date" class="excel-date-input" value="${formattedDate}">
        </td>
        <td>${new Date().toLocaleString()}</td>
        <td>
            <button class="action-btn delete-btn" onclick="deleteExcelRow(this)">Delete</button>
        </td>
    `;
    
    tbody.appendChild(tr);
}

// Function to delete a row from the Excel table
function deleteExcelRow(button) {
    const row = button.closest('tr');
    row.remove();
}

// Function to save Excel data
function saveExcelData() {
    const tbody = document.getElementById('excelTableBody');
    const rows = tbody.querySelectorAll('tr');
    
    if (rows.length === 0) {
        alert('No data to save.');
        return;
    }
    
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    
    // Create array for worksheet data
    const worksheetData = [];
    
    // Add header row
    worksheetData.push(['Supplier', 'Stock Code', 'Mass (kg)', 'Date', 'Timestamp']);
    
    // Add data rows
    rows.forEach(row => {
        const supplierSelect = row.querySelector('.excel-supplier-select');
        const stockCodeSelect = row.querySelector('.excel-stock-code-select');
        const massInput = row.querySelector('.excel-mass-input');
        const dateInput = row.querySelector('.excel-date-input');
        const timestamp = row.cells[4].textContent;
        
        const supplier = supplierSelect.value;
        const stockCode = stockCodeSelect.value;
        const mass = massInput.value || '0';
        const date = dateInput.value;
        const formattedDate = formatDate(date);
        
        worksheetData.push([supplier, stockCode, mass, formattedDate, timestamp]);
    });
    
    // Create worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Receiving Schedule');
    
    // Save workbook to localStorage
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const base64 = btoa(String.fromCharCode.apply(null, new Uint8Array(wbout)));
    localStorage.setItem('receivingScheduleWorkbook', base64);
    
    // Create a Blob with the Excel data for download
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    
    // Create a download link
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    // Set link properties
    link.setAttribute('href', url);
    link.setAttribute('download', 'Receiving Schedule.xlsx');
    link.style.visibility = 'hidden';
    
    // Add to document, trigger download, and clean up
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    alert('Schedule data saved successfully!');
    
    // Ensure the Receiving schedule folder exists
    ensureFolderExists('Receiving schedule');
    
    // Save the file to the project folder (if running locally)
    saveFileToProjectFolder('Receiving schedule/Receiving Schedule.xlsx', blob);
}