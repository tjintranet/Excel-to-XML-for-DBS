'use strict';

document.addEventListener('DOMContentLoaded', function() {
    // DOM elements
    const excelFile = document.getElementById('excelFile');
    const downloadAllBtn = document.getElementById('downloadAllBtn');
    const clearAllBtn = document.getElementById('clearAllBtn');
    const previewTable = document.getElementById('previewTable');
    const editTable = document.getElementById('editTable');
    const previewTableSection = document.getElementById('previewTableSection');
    const editTableSection = document.getElementById('editTableSection');
    const previewMode = document.getElementById('previewMode');
    const editMode = document.getElementById('editMode');
    const saveChangesBtn = document.getElementById('saveChangesBtn');
    const processingIndicator = document.getElementById('processingIndicator');
    const adjustmentSummary = document.getElementById('adjustmentSummary');
    
    // Data storage
    let xmlDataArray = [];
    let originalHeaders = [];
    let rawData = [];
    let adjustedRows = new Set();
    
    // Columns to show in edit mode
    const editableColumns = [
        'Wi_Number', 'Limp_ISBN', 'Cased_ISBN', 'Title',
        'Trim_Height', 'Trim_Width', 'Page_Extent', 'Spine_Size',
        'Reel_Width', 'Cut_Off', 'Imposition', 'Paper_Code'
    ];
    
    // Fixed trim off head value (3mm = 0030)
    function getTrimOffHeadValue() {
        return "0030";
    }
    
    // Function to pad a number with leading zeros to a specific length
    function padWithZeros(num, targetLength) {
        return num.toString().padStart(targetLength, '0');
    }
    
    // Function to copy barcode string to clipboard
    function copyBarcodeToClipboard(barcodeString, element) {
        if (!barcodeString) {
            alert('No barcode string available to copy');
            return;
        }
        
        navigator.clipboard.writeText(barcodeString).then(function() {
            // Show temporary feedback
            const originalText = element.textContent;
            const originalColor = element.style.color;
            
            element.textContent = 'Copied!';
            element.style.color = '#000000ff';
            element.style.fontWeight = 'bold';
            
            setTimeout(() => {
                element.textContent = originalText;
                element.style.color = originalColor;
                element.style.fontWeight = '';
            }, 1000);
        }).catch(function(err) {
            // Fallback for older browsers
            const textArea = document.createElement('textarea');
            textArea.value = barcodeString;
            document.body.appendChild(textArea);
            textArea.select();
            try {
                document.execCommand('copy');
                // Show feedback for fallback method
                const originalText = element.textContent;
                element.textContent = 'Copied!';
                element.style.color = '#28a745';
                setTimeout(() => {
                    element.textContent = originalText;
                    element.style.color = '';
                }, 1000);
            } catch (err) {
                alert('Failed to copy barcode string');
            }
            document.body.removeChild(textArea);
        });
    }
    
    // Function to generate DataMatrix barcode string from Excel data
    function generateBarcodeString(row) {
        try {
            const limpIsbnIndex = originalHeaders.indexOf('Limp_ISBN');
            const casedIsbnIndex = originalHeaders.indexOf('Cased_ISBN');
            const heightIndex = originalHeaders.indexOf('Trim_Height');
            const widthIndex = originalHeaders.indexOf('Trim_Width');
            const spineIndex = originalHeaders.indexOf('Spine_Size');
            const cutOffIndex = originalHeaders.indexOf('Cut_Off');
            
            // Get ISBN and determine transfer station
            let isbn = '';
            let transferStation = '1';
            
            const limpIsbnValue = limpIsbnIndex !== -1 ? row[limpIsbnIndex] : '';
            const casedIsbnValue = casedIsbnIndex !== -1 ? row[casedIsbnIndex] : '';
            
            if (limpIsbnValue && limpIsbnValue.trim() !== '') {
                isbn = limpIsbnValue;
                transferStation = '1';
            } else if (casedIsbnValue && casedIsbnValue.trim() !== '') {
                isbn = casedIsbnValue;
                transferStation = '2';
            } else {
                return null;
            }
            
            // Clean and format ISBN
            isbn = String(isbn).replace(/\D/g, '');
            if (isbn.length > 13) isbn = isbn.substring(0, 13);
            else if (isbn.length < 13) isbn = isbn.padStart(13, '0');
            
            // Get other values
            const height = parseFloat(row[heightIndex]) || 0;
            const width = parseFloat(row[widthIndex]) || 0;
            const spine = parseFloat(row[spineIndex]) || 0;
            const cutOff = parseFloat(row[cutOffIndex]) || 0;
            
            // Format barcode string (37 digits total)
            const endsheetHeight = "0000";
            
            // Spine formatting
            const spineRounded = Math.round(spine);
            let spineSegment;
            if (spineRounded >= 10) {
                spineSegment = spineRounded.toString() + "0";
            } else {
                spineSegment = "0" + spineRounded.toString() + "0";
            }
            
            const bbHeightSegment = padWithZeros(Math.round(cutOff), 3) + "0";
            const trimOffSegment = getTrimOffHeadValue();
            const trimHeightSegment = padWithZeros(Math.round(height), 3) + "0";
            const trimWidthSegment = padWithZeros(Math.round(width), 3) + "0";
            
            const barcodeString = 
                isbn + endsheetHeight + spineSegment + bbHeightSegment + 
                trimOffSegment + trimHeightSegment + trimWidthSegment + transferStation;
            
            if (barcodeString.length !== 37 || /\D/.test(barcodeString)) {
                return null;
            }
            
            return barcodeString;
        } catch (error) {
            console.error("Error generating barcode:", error);
            return null;
        }
    }
    
    function showProcessingIndicator(text) {
        if (processingIndicator) {
            document.getElementById('processingText').textContent = text;
            processingIndicator.style.display = 'block';
        }
    }

    function hideProcessingIndicator() {
        if (processingIndicator) {
            processingIndicator.style.display = 'none';
        }
    }

    function showAdjustmentSummary(count) {
        if (adjustmentSummary) {
            const adjustmentText = document.getElementById('adjustmentText');
            adjustmentText.textContent = `${count} row(s) with "Limp P/Bound 8pp Cover" production route detected. Trim_Width values have been automatically increased by 10.`;
            adjustmentSummary.classList.remove('d-none');
        }
    }

    function hideAdjustmentSummary() {
        if (adjustmentSummary) {
            adjustmentSummary.classList.add('d-none');
        }
    }

    function createXmlFromRow(headers, values) {
        const xmlDoc = document.implementation.createDocument(null, "csv", null);
        const root = xmlDoc.documentElement;
        
        const pi = xmlDoc.createProcessingInstruction('xml', 'version="1.0" encoding="UTF-8"');
        xmlDoc.insertBefore(pi, root);
        
        // Add all the regular data elements first
        for (let i = 0; i < headers.length; i++) {
            if (headers[i]) {
                const elementName = headers[i].replace(/\s+/g, '_');
                const element = xmlDoc.createElement(elementName);
                
                let value = values[i] !== undefined ? values[i].toString().trim() : '';
                
                if (elementName === 'Title') {
                    value = value.replace(/,/g, '');
                }
                
                element.textContent = value;
                root.appendChild(element);
            }
        }
        
        // Generate and add barcode string as the last element
        const barcodeString = generateBarcodeString(values);
        if (barcodeString) {
            const barcodeElement = xmlDoc.createElement('Barcode');
            barcodeElement.textContent = barcodeString;
            root.appendChild(barcodeElement);
        }
        
        const serializer = new XMLSerializer();
        let xmlString = serializer.serializeToString(xmlDoc);
        
        xmlString = xmlString.replace(/></g, '>\n<');
        xmlString = xmlString.replace(/<csv>/, '<csv>\n');
        xmlString = xmlString.replace(/<\/csv>/, '\n</csv>');
        xmlString = xmlString.replace(/<([^/?][^>]*)>/g, '  <$1>');
        
        return xmlString;
    }

    function applyProductionRouteLogic() {
        showProcessingIndicator('Applying production route logic...');
        
        const productionRouteIndex = originalHeaders.indexOf('Production_Route');
        const trimWidthIndex = originalHeaders.indexOf('Trim_Width');
        
        if (productionRouteIndex === -1 || trimWidthIndex === -1) {
            hideProcessingIndicator();
            return 0;
        }
        
        let adjustmentCount = 0;
        adjustedRows.clear();
        
        for (let i = 0; i < rawData.length; i++) {
            const row = rawData[i];
            const productionRoute = row[productionRouteIndex];
            const currentTrimWidth = row[trimWidthIndex];
            
            if (productionRoute && productionRoute.toString().trim() === 'Limp P/Bound 8pp Cover') {
                let trimWidthValue = parseFloat(currentTrimWidth) || 0;
                trimWidthValue += 10;
                rawData[i][trimWidthIndex] = trimWidthValue.toString();
                adjustedRows.add(i);
                adjustmentCount++;
            }
        }
        
        hideProcessingIndicator();
        
        if (adjustmentCount > 0) {
            showAdjustmentSummary(adjustmentCount);
        } else {
            hideAdjustmentSummary();
        }
        
        return adjustmentCount;
    }

    function regenerateAllData() {
        showProcessingIndicator('Regenerating XML data...');
        
        xmlDataArray = [];
        
        for (let i = 0; i < rawData.length; i++) {
            const values = rawData[i];
            if (values.length === 0) continue;
            
            const xml = createXmlFromRow(originalHeaders, values);
            const barcodeString = generateBarcodeString(values);
            
            const wiNumberIndex = originalHeaders.indexOf('Wi_Number');
            const limpIsbnIndex = originalHeaders.indexOf('Limp_ISBN');
            const casedIsbnIndex = originalHeaders.indexOf('Cased_ISBN');
            const titleIndex = originalHeaders.indexOf('Title');
            const productionRouteIndex = originalHeaders.indexOf('Production_Route');
            const trimWidthIndex = originalHeaders.indexOf('Trim_Width');
            
            const wiNumber = values[wiNumberIndex] || 'unknown';
            const limpIsbn = values[limpIsbnIndex] || '';
            const casedIsbn = values[casedIsbnIndex] || '';
            const isbn = limpIsbn || casedIsbn || 'N/A';
            
            let title = values[titleIndex] || 'N/A';
            if (title !== 'N/A') {
                title = title.toString().replace(/,/g, '');
            }
            
            const productionRoute = values[productionRouteIndex] || 'N/A';
            const trimWidth = values[trimWidthIndex] || 'N/A';
            
            xmlDataArray.push({
                wiNumber: wiNumber,
                isbn: isbn,
                title: title,
                productionRoute: productionRoute,
                trimWidth: trimWidth,
                barcodeString: barcodeString,
                xml: xml,
                rowIndex: i
            });
        }
        
        hideProcessingIndicator();
    }

    function downloadSingleXml(xml, wiNumber) {
        const blob = new Blob([xml], { type: 'application/xml' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${wiNumber}.xml`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    }

    async function downloadAllFiles() {
        console.log('Download function called');
        showProcessingIndicator('Generating XML files...');
        
        try {
            const zip = new JSZip();
            
            // Add all XML files to ZIP
            xmlDataArray.forEach(data => {
                zip.file(`${data.wiNumber}.xml`, data.xml);
            });
            
            // Generate and download the ZIP
            const content = await zip.generateAsync({ type: "blob" });
            const url = window.URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            
            // Create filename based on number of records
            let zipFilename = "xml_files.zip";
            if (xmlDataArray.length > 1) {
                // Multiple records: use "multi" prefix
                zipFilename = "multi_xml_files.zip";
            } else if (xmlDataArray.length === 1) {
                // Single record: use Wi Number prefix
                if (xmlDataArray[0].wiNumber && xmlDataArray[0].wiNumber !== 'unknown') {
                    zipFilename = `${xmlDataArray[0].wiNumber}_xml_files.zip`;
                }
            }
            
            a.download = zipFilename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
        } catch (error) {
            console.error('Error generating ZIP:', error);
            alert('Error generating files: ' + error.message);
        } finally {
            hideProcessingIndicator();
        }
    }

    function updatePreviewTable() {
        const tbody = previewTable.querySelector('tbody');
        tbody.innerHTML = '';
        
        xmlDataArray.forEach((xmlData, index) => {
            const row = document.createElement('tr');
            row.dataset.originalIndex = index;
            
            if (adjustedRows.has(index)) {
                row.classList.add('trim-width-adjusted');
            }
            
            // Wi Number cell
            const wiNumberCell = document.createElement('td');
            wiNumberCell.textContent = xmlData.wiNumber;
            row.appendChild(wiNumberCell);
            
            // ISBN cell
            const isbnCell = document.createElement('td');
            isbnCell.textContent = xmlData.isbn;
            row.appendChild(isbnCell);
            
            // Title cell
            const titleCell = document.createElement('td');
            titleCell.textContent = xmlData.title;
            row.appendChild(titleCell);
            
            // Production Route cell
            const productionRouteCell = document.createElement('td');
            productionRouteCell.textContent = xmlData.productionRoute;
            row.appendChild(productionRouteCell);
            
            // Trim Width cell
            const trimWidthCell = document.createElement('td');
            trimWidthCell.textContent = xmlData.trimWidth;
            if (adjustedRows.has(index)) {
                trimWidthCell.innerHTML = `<strong>${xmlData.trimWidth}</strong> <i class="bi bi-plus-circle-fill text-warning" title="Adjusted +10"></i>`;
            }
            row.appendChild(trimWidthCell);
            
            // Barcode String cell
            const barcodeCell = document.createElement('td');
            barcodeCell.className = 'barcode-string';
            if (xmlData.barcodeString) {
                barcodeCell.textContent = xmlData.barcodeString;
                barcodeCell.title = 'Click to copy barcode string to clipboard';
                barcodeCell.style.cursor = 'pointer';
                barcodeCell.onclick = () => copyBarcodeToClipboard(xmlData.barcodeString, barcodeCell);
            } else {
                barcodeCell.textContent = 'N/A';
                barcodeCell.style.color = '#6c757d';
            }
            row.appendChild(barcodeCell);
            
            // XML cell (renamed from BC Data)
            const actionCell = document.createElement('td');
            actionCell.className = 'text-center';
            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'btn btn-primary btn-sm';
            downloadBtn.innerHTML = '<i class="bi bi-download"></i>';
            downloadBtn.title = 'Download XML';
            downloadBtn.onclick = () => downloadSingleXml(xmlData.xml, xmlData.wiNumber);
            actionCell.appendChild(downloadBtn);
            row.appendChild(actionCell);
            
            tbody.appendChild(row);
        });
    }

    function populateEditTable() {
        if (!editTable) return;
        
        const thead = editTable.querySelector('thead tr');
        const tbody = editTable.querySelector('tbody');
        thead.innerHTML = '';
        tbody.innerHTML = '';
        
        const columnIndices = {};
        editableColumns.forEach(column => {
            const index = originalHeaders.indexOf(column);
            if (index !== -1) {
                columnIndices[column] = index;
            }
        });
        
        // Add headers for columns that exist in the data + Barcode String column
        editableColumns.forEach(column => {
            if (columnIndices[column] !== undefined) {
                const th = document.createElement('th');
                th.textContent = column;
                th.setAttribute('scope', 'col');
                thead.appendChild(th);
            }
        });
        
        // Add Barcode String column header
        const barcodeHeader = document.createElement('th');
        barcodeHeader.textContent = 'Barcode String';
        barcodeHeader.style.minWidth = '200px';
        thead.appendChild(barcodeHeader);
        
        // Add rows with editable cells
        rawData.forEach((rowData, rowIndex) => {
            const row = document.createElement('tr');
            row.dataset.rowIndex = rowIndex;
            
            if (adjustedRows.has(rowIndex)) {
                row.classList.add('trim-width-adjusted');
            }
            
            editableColumns.forEach(column => {
                const colIndex = columnIndices[column];
                
                if (colIndex !== undefined) {
                    const cell = document.createElement('td');
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.className = 'form-control';
                    input.value = rowData[colIndex] !== undefined ? rowData[colIndex] : '';
                    input.dataset.colIndex = colIndex;
                    
                    if (input.value.length > 20) {
                        input.title = input.value;
                    }
                    
                    cell.appendChild(input);
                    row.appendChild(cell);
                }
            });
            
            // Add Barcode String cell
            const barcodeCell = document.createElement('td');
            barcodeCell.className = 'barcode-string';
            barcodeCell.style.padding = '8px';
            
            // Generate barcode for this row
            const barcodeString = generateBarcodeString(rowData);
            if (barcodeString) {
                barcodeCell.textContent = barcodeString;
                barcodeCell.title = 'Click to copy barcode string to clipboard';
                barcodeCell.style.cursor = 'pointer';
                barcodeCell.onclick = () => copyBarcodeToClipboard(barcodeString, barcodeCell);
            } else {
                barcodeCell.textContent = 'N/A';
                barcodeCell.style.color = '#6c757d';
                barcodeCell.title = 'Cannot generate barcode - missing required data';
            }
            
            row.appendChild(barcodeCell);
            tbody.appendChild(row);
        });
        
        const rowCount = rawData.length;
        const rowCountInfo = document.getElementById('rowCountInfo');
        if (rowCountInfo) {
            rowCountInfo.textContent = `${rowCount} record${rowCount !== 1 ? 's' : ''} found`;
        }
    }

    function updateDataFromEditTable() {
        const tbody = editTable.querySelector('tbody');
        const rows = tbody.querySelectorAll('tr');
        
        rows.forEach((row, rowIndex) => {
            const inputs = row.querySelectorAll('input');
            inputs.forEach(input => {
                const colIndex = parseInt(input.dataset.colIndex);
                rawData[rowIndex][colIndex] = input.value;
            });
        });
    }

    // Function to parse XML file and extract data
    function parseXmlFile(xmlContent) {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
            
            // Check for parsing errors
            const parserError = xmlDoc.querySelector('parsererror');
            if (parserError) {
                throw new Error('Invalid XML format');
            }
            
            const csvElement = xmlDoc.querySelector('csv');
            if (!csvElement) {
                throw new Error('XML does not contain expected csv root element');
            }
            
            // Extract all child elements
            const elements = csvElement.children;
            const xmlData = {};
            
            for (let i = 0; i < elements.length; i++) {
                const element = elements[i];
                const tagName = element.tagName;
                const textContent = element.textContent.trim();
                xmlData[tagName] = textContent;
            }
            
            return xmlData;
        } catch (error) {
            console.error('Error parsing XML:', error);
            throw error;
        }
    }
    
    // Function to convert XML data back to Excel-like structure
    function convertXmlToExcelFormat(xmlDataList) {
        if (xmlDataList.length === 0) return { headers: [], data: [] };
        
        // Get all unique keys from all XML objects to create headers
        const allKeys = new Set();
        xmlDataList.forEach(xmlData => {
            Object.keys(xmlData).forEach(key => {
                if (key !== 'Barcode') { // Exclude barcode from headers as it's generated
                    allKeys.add(key);
                }
            });
        });
        
        const headers = Array.from(allKeys);
        
        // Convert XML data to rows
        const data = xmlDataList.map(xmlData => {
            return headers.map(header => xmlData[header] || '');
        });
        
        return { headers, data };
    }
    
    // Function to process uploaded XML files
    function processXmlFiles(files) {
        showProcessingIndicator('Reading XML files...');
        
        const xmlDataList = [];
        let filesProcessed = 0;
        
        const processFile = (file) => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const xmlContent = e.target.result;
                        const xmlData = parseXmlFile(xmlContent);
                        xmlDataList.push(xmlData);
                        resolve();
                    } catch (error) {
                        reject(new Error(`Error processing ${file.name}: ${error.message}`));
                    }
                };
                reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
                reader.readAsText(file);
            });
        };
        
        // Process all files
        const filePromises = Array.from(files).map(processFile);
        
        Promise.all(filePromises)
            .then(() => {
                showProcessingIndicator('Converting XML data to editable format...');
                
                // Convert XML data to Excel-like format
                const { headers, data } = convertXmlToExcelFormat(xmlDataList);
                
                originalHeaders = headers;
                rawData = data;
                
                // Reset adjusted rows since we're loading from XML
                adjustedRows.clear();
                
                showProcessingIndicator('Generating preview data...');
                
                // Generate XML and preview data
                xmlDataArray = [];
                regenerateAllData();
                
                showProcessingIndicator('Updating interface...');
                
                // Update UI
                updatePreviewTable();
                populateEditTable();
                
                // Enable buttons
                downloadAllBtn.disabled = xmlDataArray.length === 0;
                clearAllBtn.disabled = xmlDataArray.length === 0;
                
                hideProcessingIndicator();
                
                console.log(`Successfully loaded ${xmlDataList.length} XML file(s)`);
                
                // Show success message
                if (adjustmentSummary) {
                    const adjustmentText = document.getElementById('adjustmentText');
                    adjustmentText.textContent = `Successfully loaded ${xmlDataList.length} XML file(s). Data is now available for editing and re-export.`;
                    adjustmentSummary.classList.remove('d-none');
                    adjustmentSummary.className = 'alert alert-success';
                    
                    // Reset to warning class after 5 seconds
                    setTimeout(() => {
                        adjustmentSummary.className = 'alert alert-warning d-none';
                    }, 5000);
                }
            })
            .catch((error) => {
                hideProcessingIndicator();
                console.error('Error processing XML files:', error);
                alert(`Error processing XML files: ${error.message}`);
            });
    }

    function convertExcelToXml(file) {
        showProcessingIndicator('Reading Excel file...');
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            showProcessingIndicator('Parsing Excel data...');
            
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array',
                cellDates: true,
                cellNF: true,
                cellText: true,
                raw: true
            });

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            originalHeaders = jsonData[0];
            rawData = jsonData.slice(1).filter(row => row.length > 0);
            
            const adjustmentCount = applyProductionRouteLogic();
            
            showProcessingIndicator('Generating XML data...');
            
            xmlDataArray = [];
            regenerateAllData();
            
            showProcessingIndicator('Updating interface...');
            
            updatePreviewTable();
            populateEditTable();
            
            if (downloadAllBtn) downloadAllBtn.disabled = xmlDataArray.length === 0;
            if (clearAllBtn) clearAllBtn.disabled = xmlDataArray.length === 0;
            
            hideProcessingIndicator();
            
            console.log(`Processing complete. ${adjustmentCount} rows adjusted for Production Route logic.`);
        };

        reader.readAsArrayBuffer(file);
    }

    function clearAll() {
        // Don't clear the file input value - let it show the current file
        if (downloadAllBtn) downloadAllBtn.disabled = true;
        if (clearAllBtn) clearAllBtn.disabled = true;
        if (previewTable) previewTable.querySelector('tbody').innerHTML = '';
        if (editTable) {
            editTable.querySelector('thead tr').innerHTML = '';
            editTable.querySelector('tbody').innerHTML = '';
        }
        xmlDataArray = [];
        originalHeaders = [];
        rawData = [];
        adjustedRows.clear();
        
        hideProcessingIndicator();
        hideAdjustmentSummary();
        
        if (previewMode) previewMode.checked = true;
        if (previewTableSection) previewTableSection.style.display = 'block';
        if (editTableSection) editTableSection.style.display = 'none';
    }

    function clearAllComplete() {
        // This function clears everything including the file input
        if (excelFile) excelFile.value = '';
        clearAll();
    }

    // Mode Toggle Buttons
    if (previewMode) {
        previewMode.addEventListener('change', function() {
            if(this.checked) {
                previewTableSection.style.display = 'block';
                editTableSection.style.display = 'none';
            }
        });
    }
    
    if (editMode) {
        editMode.addEventListener('change', function() {
            if(this.checked) {
                previewTableSection.style.display = 'none';
                editTableSection.style.display = 'block';
            }
        });
    }

    // Save changes from edit mode
    if (saveChangesBtn) {
        saveChangesBtn.addEventListener('click', function() {
            updateDataFromEditTable();
            regenerateAllData();
            updatePreviewTable();
            previewMode.checked = true;
            previewTableSection.style.display = 'block';
            editTableSection.style.display = 'none';
        });
    }

    // Event Listeners
    if (excelFile) {
        excelFile.addEventListener('change', function() {
            const files = this.files;
            if (files && files.length > 0) {
                // Clear existing data (but don't clear file input)
                clearAll();
                
                // Check file types
                const xmlFiles = Array.from(files).filter(file => 
                    file.name.toLowerCase().endsWith('.xml')
                );
                const excelFiles = Array.from(files).filter(file => 
                    file.name.toLowerCase().endsWith('.xlsx') || 
                    file.name.toLowerCase().endsWith('.xls')
                );
                
                if (xmlFiles.length > 0 && excelFiles.length > 0) {
                    alert('Please select either Excel files OR XML files, not both types together.');
                    clearAllComplete();
                    return;
                }
                
                if (xmlFiles.length > 0) {
                    // Process XML files
                    processXmlFiles(xmlFiles);
                } else if (excelFiles.length === 1) {
                    // Process single Excel file
                    convertExcelToXml(excelFiles[0]);
                } else if (excelFiles.length > 1) {
                    alert('Please select only one Excel file at a time.');
                    clearAllComplete();
                    return;
                } else {
                    alert('Please select valid Excel (.xlsx, .xls) or XML (.xml) files.');
                    clearAllComplete();
                    return;
                }
            } else {
                clearAllComplete();
            }
        });
    }

    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', clearAllComplete);
    }
    
    if (downloadAllBtn) {
        downloadAllBtn.addEventListener('click', downloadAllFiles);
    }
    
    console.log('Script loaded successfully');
});