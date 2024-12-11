Excel File Processor
Table of Contents
Introduction
Features
Technologies Used
Prerequisites
Installation
Usage
Project Structure
Configuration
Contributing
License
Introduction
The Excel File Processor is a web application designed to upload, process, and analyze Excel files containing company data. It fetches additional information from external sources, normalizes company names, and categorizes data based on predefined codes. The processed data is then provided as downloadable Excel files, packaged in a ZIP archive.

Features
Excel File Upload: Users can upload .xlsx or .xls files through a user-friendly interface.
Data Processing: The server processes the uploaded data, normalizes company names, and fetches additional information such as EMTAK codes and employee counts.
Data Categorization: Categorizes companies based on EMTAK codes into matching and non-matching categories.
Downloadable Results: Provides processed data as downloadable Excel files within a ZIP archive.
Responsive UI: Clean and responsive user interface using Sakura CSS for styling.
Technologies Used
Frontend:

HTML5
CSS3 (Sakura CSS)
JavaScript
SheetJS (xlsx) for Excel file handling
JSZip for ZIP file generation
Backend:

Node.js
Express.js
ExcelJS for Excel file creation
Cheerio for HTML parsing
Puppeteer for web scraping
PDFKit for PDF generation
Prerequisites
Node.js (v14 or later)
npm (v6 or later)
Installation
Clone the Repository


git clone https://github.com/your-username/excel-file-processor.git
cd excel-file-processor
Install Dependencies


npm install
The project relies on the following main dependencies:

express
body-parser
node-fetch
cheerio
exceljs
pdfkit
puppeteer
jszip
Set Up Environment Variables

If your project requires environment variables, create a .env file in the root directory and add the necessary configurations. (Note: The provided code does not specify environment variables, but this is a common practice.)

Usage
Start the Server


node server.js
The server will start on http://localhost:3000.

Access the Application

Open your web browser and navigate to http://localhost:3000. You will see an interface to upload your Excel file.

Upload and Process File

Click on the "Choose File" button to select your .xlsx or .xls file.
Click the "Upload" button to submit the file.
The application will process the file, fetch additional data, and provide a downloadable ZIP archive containing the processed Excel files.
Project Structure
java
Copy code
excel-file-processor/
├── public/
│   ├── index.html
│   ├── script.js
│   └── styles.css
├── files/
│   └── (Generated files will be stored here)
├── countyMap.js
├── emtakCodes.js
├── server.js
├── package.json
├── package-lock.json
└── README.md
public/: Contains frontend assets like HTML, CSS, and JavaScript files.
files/: Directory for storing generated files (e.g., PDFs, Excel files).
countyMap.js: JavaScript module mapping county names.
emtakCodes.js: JavaScript module containing EMTAK codes.
server.js: Main server-side application file.
package.json: Lists project dependencies and scripts.
README.md: Project documentation.
Configuration
countyMap.js
Maps county names from full names to their shortened versions.


const countyMap = {
    "Harju maakond": "Harjumaa",
    // ... other mappings
};
export default countyMap;
emtakCodes.js
Contains a list of valid EMTAK codes used for categorizing companies.


const emtakCodes = [
    "1011", "10111", "1012", // ... other codes
];
export default emtakCodes;
server.js
Handles server-side processing, including file uploads, data normalization, web scraping, and Excel file generation.


import express from 'express';
import bodyParser from 'body-parser';
// ... other imports

const app = express();
const port = 3000;

// Middleware setup
app.use(bodyParser.json());
app.use(express.static(__dirname));
// ... other configurations

app.post('/upload', async (req, res) => {
    // Data processing logic
});

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});