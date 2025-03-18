let excelData = [];
let cutoffMap = {};

// Read Excel File
document.getElementById('fileInput').addEventListener('change', function(event) {
    let file = event.target.files[0];
    let reader = new FileReader();
    
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        excelData = XLSX.utils.sheet_to_json(sheet);

        displayTable();
        generateControls();
    };
    
    reader.readAsArrayBuffer(file);
});

// Display Table
function displayTable() {
    let table = document.getElementById('dataTable');
    table.innerHTML = "";

    if (excelData.length === 0) return;

    // Create Table Header
    let header = table.createTHead();
    let headerRow = header.insertRow();
    Object.keys(excelData[0]).forEach(key => {
        let th = document.createElement("th");
        th.innerText = key;
        headerRow.appendChild(th);
    });

    // Create Table Body
    let body = table.createTBody();
    excelData.forEach(rowData => {
        let row = body.insertRow();
        Object.keys(excelData[0]).forEach(key => {
            let cell = row.insertCell();
            cell.innerText = rowData[key] !== undefined ? rowData[key] : "";
        });
    });
}

// Generate Cutoff Controls
function generateControls() {
    let uniqueDegrees = [...new Set(excelData.map(row => row["Final Degree"]))];
    let controlsDiv = document.getElementById('controls');
    controlsDiv.innerHTML = "<h3>Set Cutoff Values</h3>";

    uniqueDegrees.forEach(degree => {
        let label = document.createElement("label");
        label.innerText = `Cutoff for ${degree}: `;

        let input = document.createElement("input");
        input.type = "number";
        input.id = `cutoff_${degree}`;
        input.value = 0;

        controlsDiv.appendChild(label);
        controlsDiv.appendChild(input);
        controlsDiv.appendChild(document.createElement("br"));
    });

    // Add "Apply Cutoff" Button
    let applyButton = document.createElement("button");
    applyButton.innerText = "Apply Cutoff";
    applyButton.onclick = applyCutoff;
    controlsDiv.appendChild(applyButton);
}

// Apply Cutoff Logic
function applyCutoff() {
    let uniqueDegrees = [...new Set(excelData.map(row => row["Final Degree"]))];

    // Read cutoff values
    uniqueDegrees.forEach(degree => {
        let inputElement = document.getElementById(`cutoff_${degree}`);
        if (inputElement) {
            cutoffMap[degree] = parseFloat(inputElement.value) || 0;
        }
    });

    // Counters for statistics
    let rejectedCount = 0;
    let selectedCount = 0;
    let absenteeCount = 0;

    // Update Status Based on Cutoff
    excelData = excelData.map(row => {
        let finalDegree = row["Final Degree"];
        let cutoffValue = cutoffMap[finalDegree] || 0;
        
        // Check both score fields
        let score75 = parseFloat(row["Overall Score (Max. 75)"]) || 0;
        let score60 = parseFloat(row["Overall Score (Max. 60)"]) || 0;
        
        // Use the appropriate score based on which one is available (non-zero)
        let effectiveScore = score75 > 0 ? score75 : score60;
        
        // If both scores are available, use the appropriate one based on the degree or other logic
        // This can be customized based on your specific requirements
        
        // Apply cutoff logic
        if (effectiveScore > 0 && cutoffValue > 0) {
            row["Status"] = effectiveScore < cutoffValue ? "R" : "P";
        } else {
            // If no valid scores or cutoff, set as "A" (Absent)
            row["Status"] = "A";
        }

        // Count statistics
        if (row["Status"] === "R") rejectedCount++;
        if (row["Status"] === "P") selectedCount++;
        if (row["Status"] === "A") absenteeCount++;

        return row;
    });

    displayTable(); // Refresh Table to Reflect Updates
    updateStatistics(rejectedCount, selectedCount, absenteeCount); // Update Stats Display
}

// Update Statistics Display
function updateStatistics(rejected, selected, absentees) {
    let statsDiv = document.getElementById('statistics');
    statsDiv.innerHTML = `
        <h3>Statistics</h3>
        <p><strong>No. of Rejected:</strong> ${rejected}</p>
        <p><strong>No. of Selected:</strong> ${selected}</p>
        <p><strong>No. of Absentees:</strong> ${absentees}</p>
    `;

    // Add Download Button
    let downloadButton = document.createElement("button");
    downloadButton.innerText = "Download Excel";
    downloadButton.onclick = downloadExcel;
    statsDiv.appendChild(downloadButton);
}

// Download Updated Excel
function downloadExcel() {
    let worksheet = XLSX.utils.json_to_sheet(excelData);
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Updated Data");
    XLSX.writeFile(workbook, "Updated_Data.xlsx");
}