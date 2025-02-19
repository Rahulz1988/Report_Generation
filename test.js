// JavaScript for Excel Data Mapper
document.addEventListener('DOMContentLoaded', function() {
    // Elements
    const scoreSheet = document.getElementById('scoreSheet');
    const consolidatedSheet = document.getElementById('consolidatedSheet');
    const scoreSheetName = document.getElementById('scoreSheetName');
    const consolidatedSheetName = document.getElementById('consolidatedSheetName');
    const processBtn = document.getElementById('processBtn');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const resultsSection = document.getElementById('resultsSection');
    const errorSection = document.getElementById('errorSection');
    const totalRecords = document.getElementById('totalRecords');
    const recordsMapped = document.getElementById('recordsMapped');
    const recordsNotFound = document.getElementById('recordsNotFound');

    let scoreSheetData = null;
    let consolidatedSheetData = null;

    // Update file name displays
    scoreSheet.addEventListener('change', function() {
        scoreSheetName.textContent = this.files.length ? this.files[0].name : 'No file chosen';
        readExcelFile(this.files[0], 'score');
    });

    consolidatedSheet.addEventListener('change', function() {
        consolidatedSheetName.textContent = this.files.length ? this.files[0].name : 'No file chosen';
        readExcelFile(this.files[0], 'consolidated');
    });

    // Enable button only when both files are selected
    function checkEnableButton() {
        processBtn.disabled = !(scoreSheet.files.length && consolidatedSheet.files.length);
    }

    // Read Excel file
    function readExcelFile(file, type) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            if (type === 'score') {
                scoreSheetData = jsonData;
            } else {
                consolidatedSheetData = jsonData;
            }
            checkEnableButton();
        };
        reader.readAsArrayBuffer(file);
    }

    // Simulate processing when button is clicked
    processBtn.addEventListener('click', function() {
        progressContainer.classList.remove('hidden');
        simulateProcessing();
    });

    // Simulate processing with progress
    function simulateProcessing() {
        let progress = 0;

        const interval = setInterval(() => {
            progress += 5;
            progressBar.style.width = progress + '%';

            if (progress >= 100) {
                clearInterval(interval);
                progressContainer.classList.add('hidden');

                // Show results (real results would be determined by your processing logic)
                totalRecords.textContent = '457';
                recordsMapped.textContent = '423';
                recordsNotFound.textContent = '34';

                resultsSection.classList.remove('hidden');
            }
        }, 100);
    }

    // Download button handler
    document.getElementById('downloadBtn').addEventListener('click', function() {
        alert('Download functionality would be implemented here.');
    });
});