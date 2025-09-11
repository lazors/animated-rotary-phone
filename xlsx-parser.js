const XLSX = require('xlsx');
const fs = require('fs');

function parseXlsxToJson(filePath, outputPath = null) {
    try {
        const workbook = XLSX.readFile(filePath);
        const result = {};
        
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            result[sheetName] = jsonData;
        });
        
        if (outputPath) {
            fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
            console.log(`JSON data saved to: ${outputPath}`);
        }
        
        return result;
    } catch (error) {
        console.error('Error parsing XLSX file:', error.message);
        return null;
    }
}

function parseXlsxSheet(filePath, sheetName = null) {
    try {
        const workbook = XLSX.readFile(filePath);
        const targetSheet = sheetName || workbook.SheetNames[0];
        
        if (!workbook.Sheets[targetSheet]) {
            throw new Error(`Sheet "${targetSheet}" not found`);
        }
        
        const worksheet = workbook.Sheets[targetSheet];
        return XLSX.utils.sheet_to_json(worksheet);
    } catch (error) {
        console.error('Error parsing XLSX sheet:', error.message);
        return null;
    }
}

module.exports = {
    parseXlsxToJson,
    parseXlsxSheet
};

if (require.main === module) {
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('Usage: node xlsx-parser.js <input.xlsx> [output.json]');
        console.log('Example: node xlsx-parser.js data.xlsx output.json');
        process.exit(1);
    }
    
    const inputFile = args[0];
    const outputFile = args[1];
    
    const result = parseXlsxToJson(inputFile, outputFile);
    if (result && !outputFile) {
        console.log(JSON.stringify(result, null, 2));
    }
}