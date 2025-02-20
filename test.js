document.addEventListener('DOMContentLoaded', function() {
    // Elements
    const scoreSheetInput = document.getElementById('scoreSheetInput');
    const consolidatedSheetInput = document.getElementById('consolidatedSheetInput');
    const scoreSheetName = document.getElementById('scoreSheetName');
    const consolidatedSheetName = document.getElementById('consolidatedSheetName');
    const processBtn = document.getElementById('processBtn');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const resultsSection = document.getElementById('resultsSection');
    const totalRecords = document.getElementById('totalRecords');
    const recordsMapped = document.getElementById('recordsMapped');
    const recordsNotFound = document.getElementById('recordsNotFound');
    const downloadBtn = document.getElementById('downloadBtn');
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');

    // Data storage
    let scoreSheetData = null;
    let consolidatedSheetData = null;
    let processedData = null;

    // Column mapping based on requirements
    const columnMapping = {
        // Score sheet headers to Consolidated sheet headers
        "English (Max. 15)": "Verbal Ability (Max 15)",
        "Logical Reasoning (Max. 15)": "Logical Reasoning (Max 15)",
        "Quantitative Ability (Max. 15)": "Quantitative Ability (Max 15)",
        "General Knowledge (Max. 10)": "General Knowledge (Max 10)",
        "Computer Awareness (Max. 10)": "Computer Awareness (Max 10)",
        "Sales Aptitde (Max. 10)": "Sales Aptitude (Max 10)",
        "Overall Score (Max. 75)": "Overall Score (Max 75)",
        "Status": "Status",
        "Final Degree": "Final Degree",
        "Proctoring Decision (Please remove this before Uploading)": null, // No direct mapping
        "Personality Test": "Personality Test",
        "S": "Sociability",
        "T": "Team Work",
        "CA": "Cognitive Agility",
        "R": "Resilience",
        "RO": "Result Orientation",
        "C": "Conscientiousness",
        "SO": "Service Orientation",
        "Overall Score (Max. 315)": "Overall Score (Max 315)",
        "Sociability (6)": "Sociability (6)",
        "Team work (10)": "Team Work (10)",
        "Cognitive Agility (15)": "Cognitive Agility (15)",
        "Resilience (8)": "Resilience (8)",
        "Result Orientation (11)": "Result Orientation (11)",
        "Conscientiousness (7)": "Conscientiousness (7)",
        "Service Orientation (6)": "Service Orientation (6)",
        "Average": "Behavioural Average"
    };

    // Special mappings for duplicate headers
    const specialMappings = [
        // First occurrence
        { scoreHeader: "Test Status", occurrence: 1, consolidatedHeader: "Aptitude Test Status" },
        { scoreHeader: "Attended Date (DD MM YYYY)", occurrence: 1, consolidatedHeader: "Aptitude Attended Date (DD MM YYYY)" },
        { scoreHeader: "Time Spent (Duration 60 Minutes)", occurrence: 1, consolidatedHeader: "Aptitude Time Spent (Duration 60 Minutes)" },
        
        // Second occurrence
        { scoreHeader: "Test Status", occurrence: 2, consolidatedHeader: "Behavioural Test Status" },
        { scoreHeader: "Attended Date (DD MM YYYY)", occurrence: 2, consolidatedHeader: "Behavioural Attended Date (DD MM YYYY)" },
        { scoreHeader: "Time Spent (Duration 15 Minutes)", occurrence: 1, consolidatedHeader: "Behavioural Time Spent (Duration 15 Minutes)" }
    ];

    // File input handlers
    scoreSheetInput.addEventListener('change', function(e) {
        if (this.files.length > 0) {
            scoreSheetName.textContent = this.files[0].name;
            readExcelFile(this.files[0], 'score');
        } else {
            scoreSheetName.textContent = 'No file chosen';
            scoreSheetData = null;
            checkFilesLoaded();
        }
    });

    consolidatedSheetInput.addEventListener('change', function(e) {
        if (this.files.length > 0) {
            consolidatedSheetName.textContent = this.files[0].name;
            readExcelFile(this.files[0], 'consolidated');
        } else {
            consolidatedSheetName.textContent = 'No file chosen';
            consolidatedSheetData = null;
            checkFilesLoaded();
        }
    });

    // Read Excel file with better error handling and proper data type preservation
    function readExcelFile(file, type) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                if (workbook.SheetNames.length === 0) {
                    showError(`The ${type === 'score' ? 'Score Sheet' : 'Consolidated Sheet'} does not contain any worksheets.`);
                    return;
                }
                
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // Use raw: true to preserve numeric data types
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
                    header: 1, 
                    defval: "",
                    raw: true 
                });
                
                if (jsonData.length === 0) {
                    showError(`The ${type === 'score' ? 'Score Sheet' : 'Consolidated Sheet'} appears to be empty.`);
                    return;
                }
                
                if (type === 'score') {
                    scoreSheetData = jsonData;
                } else {
                    consolidatedSheetData = jsonData;
                }
                
                checkFilesLoaded();
            } catch (error) {
                showError(`Error reading ${type === 'score' ? 'Score Sheet' : 'Consolidated Sheet'}: ${error.message}`);
            }
        };
        
        reader.onerror = function() {
            showError(`Failed to read ${type === 'score' ? 'Score Sheet' : 'Consolidated Sheet'} file`);
        };
        
        reader.readAsArrayBuffer(file);
    }

    // Check if both files are loaded
    function checkFilesLoaded() {
        if (scoreSheetData && consolidatedSheetData) {
            processBtn.disabled = false;
        } else {
            processBtn.disabled = true;
        }
    }

    // Process button click handler
    processBtn.addEventListener('click', function() {
        try {
            processFiles();
        } catch (error) {
            showError("Processing error: " + error.message);
        }
    });

    // Download button click handler
    downloadBtn.addEventListener('click', function() {
        downloadProcessedSheet();
    });

    // Process the files with improved header detection
    function processFiles() {
        hideError();
        progressContainer.classList.remove('hidden');
        resultsSection.classList.add('hidden');
        
        // Get headers
        const scoreHeaders = scoreSheetData[0];
        const consolidatedHeaders = consolidatedSheetData[0];
        
        // Clean up headers to handle potential whitespace issues
        const cleanedScoreHeaders = scoreHeaders.map(h => String(h).trim());
        const cleanedConsolidatedHeaders = consolidatedHeaders.map(h => String(h).trim());
        
        // Validate basic required headers
        const headerValidation = validateRequiredHeaders(cleanedScoreHeaders, cleanedConsolidatedHeaders);
        if (!headerValidation.valid) {
            showError(headerValidation.message);
            progressContainer.classList.add('hidden');
            return;
        }
        
        // Find indices for key columns needed for matching
        const scoreIdIndex = findHeaderIndex(cleanedScoreHeaders, "Candidate Id");
        const scoreNameIndex = findHeaderIndex(cleanedScoreHeaders, "Candidate Name");
        const consolidatedIdIndex = findHeaderIndex(cleanedConsolidatedHeaders, "Candidate Id");
        const consolidatedNameIndex = findHeaderIndex(cleanedConsolidatedHeaders, "Candidate Name");
        
        if (scoreIdIndex === -1 || scoreNameIndex === -1 || 
            consolidatedIdIndex === -1 || consolidatedNameIndex === -1) {
            showError("Could not find Candidate Id or Candidate Name columns in one or both files.");
            progressContainer.classList.add('hidden');
            return;
        }
        
        // Create a deep copy of score sheet data
        processedData = JSON.parse(JSON.stringify(scoreSheetData));
        
        // Build a complete mapping of indices
        const columnIndexMap = buildColumnIndexMap(cleanedScoreHeaders, cleanedConsolidatedHeaders);
        
        if (Object.keys(columnIndexMap).length === 0) {
            showError("Could not find any mappable columns between the two sheets. Please check that your column headers match the expected format.");
            progressContainer.classList.add('hidden');
            return;
        }
        
        // Debug - log the column mappings found
        console.log("Column mappings found:", columnIndexMap);
        
        // Process data
        let mappedCount = 0;
        let notFoundCount = 0;
        const totalRecordsCount = Math.max(1, processedData.length - 1); // Exclude header row, ensure at least 1
        
        // Start from row 1 (after headers)
        for (let i = 1; i < processedData.length; i++) {
            const scoreRow = processedData[i];
            
            // Get candidate ID and name, convert to string and trim
            const candidateId = scoreRow[scoreIdIndex] ? String(scoreRow[scoreIdIndex]).trim() : "";
            const candidateName = scoreRow[scoreNameIndex] ? String(scoreRow[scoreNameIndex]).trim() : "";
            
            if (!candidateId && !candidateName) {
                notFoundCount++;
                continue; // Skip empty rows
            }
            
            // Update progress
            const progress = Math.round(((i - 1) / totalRecordsCount) * 100);
            progressBar.style.width = progress + '%';
            progressText.textContent = `Processing ${i} of ${totalRecordsCount} records (${progress}%)`;
            
            // Find matching row in consolidated data
            const matchedRow = findMatchingCandidateRow(
                consolidatedSheetData,
                consolidatedIdIndex,
                consolidatedNameIndex,
                candidateId,
                candidateName
            );
            
            if (matchedRow) {
                // Map all data from consolidated to score sheet using the index map
                for (const [scoreIndex, consolidatedIndex] of Object.entries(columnIndexMap)) {
                    const scoreIdx = parseInt(scoreIndex);
                    const consolidatedIdx = parseInt(consolidatedIndex);
                    
                    // Only update if the consolidated data has a value
                    if (matchedRow[consolidatedIdx] !== undefined && matchedRow[consolidatedIdx] !== "") {
                        scoreRow[scoreIdx] = matchedRow[consolidatedIdx];
                    }
                }
                mappedCount++;
            } else {
                notFoundCount++;
            }
        }
        
        // Complete progress bar
        progressBar.style.width = '100%';
        progressText.textContent = 'Processing complete!';
        
        // Show results
        totalRecords.textContent = totalRecordsCount;
        recordsMapped.textContent = mappedCount;
        recordsNotFound.textContent = notFoundCount;
        
        // Show results section after a short delay
        setTimeout(() => {
            progressContainer.classList.add('hidden');
            resultsSection.classList.remove('hidden');
        }, 1000);
    }
    
    // Build a comprehensive mapping between score sheet column indices and consolidated sheet column indices
    function buildColumnIndexMap(scoreHeaders, consolidatedHeaders) {
        const indexMap = {};
        
        // Process regular mappings
        for (const [scoreHeader, consolidatedHeader] of Object.entries(columnMapping)) {
            if (consolidatedHeader) {
                const scoreIndices = findAllHeaderIndices(scoreHeaders, scoreHeader);
                const consolidatedIndices = findAllHeaderIndices(consolidatedHeaders, consolidatedHeader);
                
                if (scoreIndices.length > 0 && consolidatedIndices.length > 0) {
                    // For simple 1:1 mappings, just use the first occurrence
                    indexMap[scoreIndices[0]] = consolidatedIndices[0];
                }
            }
        }
        
        // Process special mappings for duplicate headers
        for (const mapping of specialMappings) {
            const scoreIndices = findAllHeaderIndices(scoreHeaders, mapping.scoreHeader);
            const consolidatedIndices = findAllHeaderIndices(consolidatedHeaders, mapping.consolidatedHeader);
            
            if (scoreIndices.length >= mapping.occurrence && consolidatedIndices.length > 0) {
                indexMap[scoreIndices[mapping.occurrence - 1]] = consolidatedIndices[0];
            }
        }
        
        return indexMap;
    }
    
    // Find a header index with case-insensitive matching and trimming
    function findHeaderIndex(headers, targetHeader) {
        const normalizedTarget = String(targetHeader).toLowerCase().trim();
        return headers.findIndex(header => 
            String(header).toLowerCase().trim() === normalizedTarget
        );
    }
    
    // Find all occurrences of a header with flexible matching
    function findAllHeaderIndices(headers, targetHeader) {
        const normalizedTarget = String(targetHeader).toLowerCase().trim();
        const indices = [];
        
        headers.forEach((header, index) => {
            const currentHeader = String(header).toLowerCase().trim();
            
            // Try exact match first
            if (currentHeader === normalizedTarget) {
                indices.push(index);
                return;
            }
            
            // Try matching without parentheses content for personality traits
            // This handles "Team work (10)" vs "Team Work (10)" type issues
            if (normalizedTarget.includes('(') && currentHeader.includes('(')) {
                const targetBase = normalizedTarget.split('(')[0].trim();
                const currentBase = currentHeader.split('(')[0].trim();
                const targetNumber = normalizedTarget.match(/\((\d+)\)/)?.[1];
                const currentNumber = currentHeader.match(/\((\d+)\)/)?.[1];
                
                if (targetBase.replace(/\s+/g, '') === currentBase.replace(/\s+/g, '') && 
                    targetNumber === currentNumber) {
                    indices.push(index);
                }
            }
            
            // Handle specific personality trait mismatches
            if (normalizedTarget === "sociability (6)" && currentHeader === "sociability(6)") {
                indices.push(index);
            }
            if (normalizedTarget === "team work (10)" && currentHeader === "teamwork(10)") {
                indices.push(index);
            }
            if (normalizedTarget === "cognitive agility (15)" && currentHeader === "cognitiveagility(15)") {
                indices.push(index);
            }
            if (normalizedTarget === "resilience (8)" && currentHeader === "resilience(8)") {
                indices.push(index);
            }
            if (normalizedTarget === "result orientation (11)" && currentHeader === "resultorientation(11)") {
                indices.push(index);
            }
            if (normalizedTarget === "conscientiousness (7)" && currentHeader === "conscientiousness(7)") {
                indices.push(index);
            }
            if (normalizedTarget === "service orientation (6)" && currentHeader === "serviceorientation(6)") {
                indices.push(index);
            }
        });
        
        return indices;
    }
    
    // Debug function to print header mappings (can be called from console)
    window.debugHeaderMappings = function() {
        if (!scoreSheetData || !consolidatedSheetData) {
            console.log("No data loaded yet");
            return;
        }
        
        const scoreHeaders = scoreSheetData[0].map(h => String(h).trim());
        const consolidatedHeaders = consolidatedSheetData[0].map(h => String(h).trim());
        
        console.log("=== SCORE SHEET HEADERS ===");
        scoreHeaders.forEach((h, i) => console.log(`${i}: "${h}"`));
        
        console.log("\n=== CONSOLIDATED SHEET HEADERS ===");
        consolidatedHeaders.forEach((h, i) => console.log(`${i}: "${h}"`));
        
        console.log("\n=== ATTEMPTED MAPPINGS ===");
        for (const [scoreHeader, consolidatedHeader] of Object.entries(columnMapping)) {
            if (consolidatedHeader) {
                const scoreIndices = findAllHeaderIndices(scoreHeaders, scoreHeader);
                const consolidatedIndices = findAllHeaderIndices(consolidatedHeaders, consolidatedHeader);
                
                console.log(`"${scoreHeader}" -> "${consolidatedHeader}"`);
                console.log(`  Score indices: ${scoreIndices.join(', ')}`);
                console.log(`  Consolidated indices: ${consolidatedIndices.join(', ')}`);
            }
        }
    };
    
    // Find matching candidate row with flexible matching strategy
    function findMatchingCandidateRow(data, idIndex, nameIndex, targetId, targetName) {
        // Skip header row
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            
            const rowId = row[idIndex] ? String(row[idIndex]).trim() : "";
            const rowName = row[nameIndex] ? String(row[nameIndex]).trim() : "";
            
            // Try exact match on both ID and name first
            if (targetId && targetName && 
                rowId === targetId && 
                rowName === targetName) {
                return row;
            }
            
            // If no exact match but we have ID, match on ID only
            if (targetId && rowId === targetId) {
                return row;
            }
            
            // Last resort, if we only have name, match on name
            if (!targetId && targetName && rowName === targetName) {
                return row;
            }
        }
        
        return null;
    }
    
    // Validate that essential headers exist
    function validateRequiredHeaders(scoreHeaders, consolidatedHeaders) {
        const requiredScoreHeaders = ["Candidate Id", "Candidate Name"];
        const requiredConsolidatedHeaders = ["Candidate Id", "Candidate Name"];
        
        // Check score sheet for required headers (case-insensitive)
        const missingScoreHeaders = requiredScoreHeaders.filter(requiredHeader => {
            const normalizedRequired = requiredHeader.toLowerCase().trim();
            return !scoreHeaders.some(header => 
                String(header).toLowerCase().trim() === normalizedRequired
            );
        });
        
        if (missingScoreHeaders.length > 0) {
            return {
                valid: false,
                message: `Missing essential headers in Score Sheet: ${missingScoreHeaders.join(", ")}`
            };
        }
        
        // Check consolidated sheet for required headers (case-insensitive)
        const missingConsolidatedHeaders = requiredConsolidatedHeaders.filter(requiredHeader => {
            const normalizedRequired = requiredHeader.toLowerCase().trim();
            return !consolidatedHeaders.some(header => 
                String(header).toLowerCase().trim() === normalizedRequired
            );
        });
        
        if (missingConsolidatedHeaders.length > 0) {
            return {
                valid: false,
                message: `Missing essential headers in Consolidated Sheet: ${missingConsolidatedHeaders.join(", ")}`
            };
        }
        
        return { valid: true };
    }
    
    // Download the processed sheet with proper numeric formatting
    function downloadProcessedSheet() {
        try {
            // Create a worksheet
            const ws = XLSX.utils.aoa_to_sheet(processedData);
            
            // Set numeric format for applicable cells
            const range = XLSX.utils.decode_range(ws['!ref']);
            
            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell_address = XLSX.utils.encode_cell({r: r, c: c});
                    const cell = ws[cell_address];
                    
                    if (cell && typeof cell.v === 'number') {
                        // Ensure numeric cells have numeric type
                        cell.t = 'n';
                        
                        // Format based on header content and value
                        const headerText = r > 0 && processedData[0][c] ? 
                            String(processedData[0][c]).toLowerCase() : '';
                        
                        // Determine appropriate format
                        if (headerText.includes('percentage') || headerText.includes('%')) {
                            cell.z = '0.00%';  // Percentage format
                        } 
                        else if (headerText.includes('score') || 
                                 headerText.includes('ability') || 
                                 headerText.includes('aptitude')) {
                            if (Number.isInteger(cell.v)) {
                                cell.z = '0';  // Integer format
                            } else {
                                cell.z = '0.00';  // Decimal format
                            }
                        }
                        else if (Number.isInteger(cell.v)) {
                            cell.z = '0';  // Integer format
                        }
                        else {
                            cell.z = '0.00';  // Default decimal format
                        }
                    }
                }
            }
            
            // Create a workbook
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Mapped Data");
            
            // Generate file and trigger download
            XLSX.writeFile(wb, "Updated_Score_Sheet.xlsx");
        } catch (error) {
            showError("Error generating file: " + error.message);
        }
    }
    
    // Show error message
    function showError(message) {
        errorMessage.textContent = message;
        errorSection.classList.remove('hidden');
        progressContainer.classList.add('hidden');
    }
    
    // Hide error message
    function hideError() {
        errorSection.classList.add('hidden');
    }
});