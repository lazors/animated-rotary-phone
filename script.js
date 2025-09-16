class XLSXParser {
    constructor() {
        this.workbook = null;
        this.currentData = null;
        this.fileName = '';
        
        this.initializeElements();
        this.attachEventListeners();
    }

    initializeElements() {
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.controls = document.getElementById('controls');
        this.sheetSelect = document.getElementById('sheetSelect');
        this.parseBtn = document.getElementById('parseBtn');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.clearBtn = document.getElementById('clearBtn');
        this.resultsSection = document.getElementById('resultsSection');
        this.dataTable = document.getElementById('dataTable');
        this.jsonOutput = document.getElementById('jsonOutput');
        this.recordCount = document.getElementById('recordCount');
        this.fileNameSpan = document.getElementById('fileName');
        this.errorMessage = document.getElementById('errorMessage');
    }

    attachEventListeners() {
        // File upload events
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));

        // Drag and drop events
        this.uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e));

        // Control buttons
        this.parseBtn.addEventListener('click', () => this.parseSelectedSheet());
        this.downloadBtn.addEventListener('click', () => this.downloadJSON());
        this.clearBtn.addEventListener('click', () => this.clearAll());

        // Error message close
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('error-close')) {
                this.hideError();
            }
        });
    }

    handleDragOver(e) {
        e.preventDefault();
        this.uploadArea.classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        this.uploadArea.classList.remove('dragover');
    }

    handleDrop(e) {
        e.preventDefault();
        this.uploadArea.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.handleFileSelect(files[0]);
        }
    }

    handleFileSelect(file) {
        if (!file) return;

        // Validate file type
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        
        if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
            this.showError('Please select a valid Excel file (.xlsx or .xls)');
            return;
        }

        this.fileName = file.name;
        this.fileNameSpan.textContent = this.fileName;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });
                this.populateSheetSelector();
                this.controls.style.display = 'flex';
                this.hideError();
            } catch (error) {
                this.showError('Error reading Excel file: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    populateSheetSelector() {
        this.sheetSelect.innerHTML = '<option value="">Choose a sheet...</option>';
        
        this.workbook.SheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            this.sheetSelect.appendChild(option);
        });

        // Auto-select first sheet if only one exists
        if (this.workbook.SheetNames.length === 1) {
            this.sheetSelect.value = this.workbook.SheetNames[0];
        }
    }

    parseSelectedSheet() {
        const selectedSheet = this.sheetSelect.value;
        if (!selectedSheet) {
            this.showError('Please select a sheet to parse');
            return;
        }

        try {
            const worksheet = this.workbook.Sheets[selectedSheet];
            this.currentData = XLSX.utils.sheet_to_json(worksheet);
            
            this.displayResults(selectedSheet);
            this.downloadBtn.style.display = 'inline-flex';
            this.hideError();
        } catch (error) {
            this.showError('Error parsing sheet: ' + error.message);
        }
    }

    displayResults(sheetName) {
        this.recordCount.textContent = `${this.currentData.length} records`;
        
        // Create table
        this.createDataTable();
        
        // Show JSON output
        this.jsonOutput.textContent = JSON.stringify(this.currentData, null, 2);
        
        this.resultsSection.style.display = 'block';
        this.resultsSection.scrollIntoView({ behavior: 'smooth' });
    }

    createDataTable() {
        if (!this.currentData || this.currentData.length === 0) {
            this.dataTable.innerHTML = '<p>No data found in the selected sheet.</p>';
            return;
        }

        // Get all unique keys from the data
        const allKeys = new Set();
        this.currentData.forEach(row => {
            Object.keys(row).forEach(key => allKeys.add(key));
        });
        const headers = Array.from(allKeys);

        // Create table
        const table = document.createElement('table');
        
        // Create header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Create body
        const tbody = document.createElement('tbody');
        this.currentData.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                const value = row[header];
                td.textContent = value !== undefined && value !== null ? value : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        this.dataTable.innerHTML = '';
        this.dataTable.appendChild(table);
    }

    downloadJSON() {
        if (!this.currentData) return;

        const jsonStr = JSON.stringify(this.currentData, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = this.fileName.replace(/\.(xlsx|xls)$/i, '.json');
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    clearAll() {
        this.workbook = null;
        this.currentData = null;
        this.fileName = '';
        
        this.fileInput.value = '';
        this.controls.style.display = 'none';
        this.resultsSection.style.display = 'none';
        this.downloadBtn.style.display = 'none';
        this.sheetSelect.innerHTML = '<option value="">Choose a sheet...</option>';
        this.hideError();
    }

    showError(message) {
        const errorText = this.errorMessage.querySelector('.error-text');
        errorText.textContent = message;
        this.errorMessage.style.display = 'block';
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            this.hideError();
        }, 5000);
    }

    hideError() {
        this.errorMessage.style.display = 'none';
    }
}

// Initialize the parser when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new XLSXParser();
});