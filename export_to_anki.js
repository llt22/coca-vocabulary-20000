const XLSX = require("xlsx");
const fs = require('fs');
const path = require('path');

// Create output directory if it doesn't exist
const outputDir = path.join(__dirname, 'anki_export');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

// CSV header
const header = "word,phonetic,definition,example,tags\n";

// Process all Excel files
fs.readdir(path.join(__dirname, 'excel'), (err, files) => {
    if (err) {
        console.error("Error reading excel directory:", err);
        return;
    }

    // Create a single CSV file for all words
    const allWordsFile = path.join(outputDir, 'all_words.csv');
    fs.writeFileSync(allWordsFile, header);
    
    // Create separate CSV files for each 1000 words
    files.forEach(file => {
        console.log(`Processing ${file}...`);
        
        // Extract part number and range from filename
        const match = file.match(/part(\d+)_(\d+)-(\d+)\.xlsx/);
        if (!match) return;
        
        const partNum = match[1];
        const startRange = match[2];
        const endRange = match[3];
        
        // Create CSV file for this part
        const outputFile = path.join(outputDir, `part${partNum}_${startRange}-${endRange}.csv`);
        fs.writeFileSync(outputFile, header);
        
        // Read Excel file
        const workbook = XLSX.readFile(path.join(__dirname, 'excel', file));
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const items = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Process each word
        items.forEach((item, index) => {
            if (!item[0]) return; // Skip empty rows
            
            const word = item[0];
            const phonetic = item[1] || '';
            const definition = item[2] || '';
            
            // Escape commas and quotes in fields
            const escapedWord = word.replace(/"/g, '""');
            const escapedPhonetic = phonetic.replace(/"/g, '""');
            const escapedDefinition = definition.replace(/"/g, '""');
            
            // Create CSV row
            const csvRow = `"${escapedWord}","${escapedPhonetic}","${escapedDefinition}","","COCA ${startRange}-${endRange}"\n`;
            
            // Append to part file
            fs.appendFileSync(outputFile, csvRow);
            
            // Append to all words file
            fs.appendFileSync(allWordsFile, csvRow);
        });
        
        console.log(`Created ${outputFile}`);
    });
    
    console.log(`\nExport complete! Files saved to ${outputDir}`);
    console.log(`\nTo import into Anki:`);
    console.log(`1. Open Anki`);
    console.log(`2. Click "Import File" from the main screen`);
    console.log(`3. Select one of the CSV files from the ${outputDir} directory`);
    console.log(`4. Make sure field mapping is correct (word, phonetic, definition, example, tags)`);
    console.log(`5. Click "Import"`);
});