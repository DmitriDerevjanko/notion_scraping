document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('excelFile');
    const uploadButton = document.getElementById('uploadButton');
    const messageDiv = document.getElementById('message');

    function uploadFile() {
        const file = fileInput.files[0];

        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header:1 });
                    fetch('/upload', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ data: sheetData }),
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.matchingFileData && data.nonMatchingFileData) {
                            downloadResults({
                                matchingFileData: data.matchingFileData,
                                nonMatchingFileData: data.nonMatchingFileData
                            });
                        } else {
                            messageDiv.textContent = 'File processed but no download data were provided.';
                        }
                    })
                    
                    .catch((error) => {
                        messageDiv.textContent = 'Error during data submission: ' + error.message;
                    });
                } catch (error) {
                    messageDiv.textContent = 'Error reading the file: ' + error.message;
                }
            };
            reader.readAsBinaryString(file);
        } else {
            messageDiv.textContent = 'Please select a file.';
        }
    }

    function downloadResults(results) {
        const zip = new JSZip();
        
        zip.file('Clients.xlsx', results.matchingFileData, {base64: true});
        zip.file('Ecosystem.xlsx', results.nonMatchingFileData, {base64: true});
        
        zip.generateAsync({type:"blob"}).then(function(content) {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = "OrganisatonsToNotion.zip";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    }
    uploadButton.addEventListener('click', uploadFile);
});
