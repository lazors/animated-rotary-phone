# XLSX Parser

A web-based Excel file parser that allows users to upload XLSX/XLS files, view their contents, and export data as JSON.

## Features

- **File Upload**: Drag and drop or browse to upload Excel files (.xlsx, .xls)
- **Sheet Selection**: Choose from multiple worksheets within an Excel file
- **Data Visualization**: View parsed data in both table and JSON format
- **JSON Export**: Download the parsed data as a JSON file
- **Responsive Design**: Clean, modern UI that works on different screen sizes

## Getting Started

### Prerequisites

- A modern web browser
- Node.js (for development dependencies)

### Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd animated-rotary-phone
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Open `index.html` in your web browser or serve it using a local server.

## Usage

1. **Upload a File**:
   - Click the upload area or drag and drop an Excel file
   - Supported formats: .xlsx, .xls

2. **Select a Sheet**:
   - Choose the worksheet you want to parse from the dropdown menu

3. **Parse Data**:
   - Click "Parse Data" to process the selected sheet
   - View the results in both table and JSON format

4. **Export**:
   - Click "Download JSON" to save the parsed data as a JSON file

## Project Structure

```
animated-rotary-phone/
├── index.html          # Main HTML file with UI structure
├── script.js           # JavaScript logic for XLSX parsing
├── styles.css          # CSS styling for the interface
├── xlsx-parser.js      # Additional XLSX parsing utilities
├── package.json        # Project dependencies
└── README.md          # This file
```

## Dependencies

- **xlsx**: Library for parsing Excel files (loaded via CDN and npm)

## Technologies Used

- HTML5
- CSS3
- Vanilla JavaScript
- SheetJS XLSX library

## Browser Support

This application works in all modern browsers that support:
- File API
- Drag and Drop API
- ES6+ JavaScript features

## License

ISC