const vecBO = [
    "9 DRD B.O",
    "Ashtapur B.O",
    "Dahitane B.O",
    "Dhanori B.O",
    "Jambe B.O",
    "Kalewadi B.O",
    "Kesnand B.O",
    "Kondhwa BK B.O",
    "Kondhwa KH B.O",
    "Kondhwa Lh B.O",
    "Koregaon Bhivar B.O",
    "Manjari BK B.O",
    "Marunji B.O",
    "Mundhva B.O",
    "Punawale B.O",
    "Rahu B.O",
    "Talawade B.O",
    "Tathawade B.O",
    "Telewadi B.O",
    "Thergaon B.O",
    "Uruli Devachi B.O",
    "Vadebholhai B.O",
    "Vadgaon Shinde B.O",
    "Vadki B.O",
    "Wakad B.O",
    "Walki B.O"
];

const vecSO = [
    "Airport S.O (Pune)",
    "Akurdi S.O",
    "Ammunition Factory Khadki S.O",
    "Aundh Camp S.O",
    "Bhosari I.E. S.O",
    "Bhosarigoan S.O",
    "Bibvewadi S.O",
    "C D A (O) S.O",
    "C M E S.O",
    "Chikhali S.O",
    "Chinchwad East S.O",
    "Chinchwadgaon S.O",
    "Dapodi Bazar S.O",
    "Dapodi S.O",
    "Dighi Camp S.O",
    "Dr.B.A. Chowk S.O",
    "Dunkirk lines S.O",
    "East Khadki S.O",
    "Ghorpuri Bazar S.O",
    "Gondhale Nagar S.O",
    "H.E. Factory S.O",
    "Hadapsar S.O",
    "Hadpsar I.E. S.O",
    "Iaf Station S.O",
    "Indrayaninagar S.O",
    "Infotech Park (Hinjawadi) S.O",
    "Kasarwadi S.O",
    "Khadki Bazar S.O",
    "Khadki S.O",
    "Lohogaon S.O",
    "M.Phulenagar S.O",
    "Maan S.O",
    "Manjari Farm S.O",
    "Market Yard S.O (Pune)",
    "Masulkar Colony S.O",
    "Mohamadwadi S.O",
    "Moshi SO",
    "Mundhva AV S.O",
    "N I B M S.O",
    "N.W. College S.O",
    "P.C.N.T. S.O",
    "Phursungi S.O",
    "Pimple Gurav S.O",
    "Pimpri Colony S.O",
    "Pimpri P F S.O",
    "Pimpri Waghire S.O",
    "Pune Cantt East S.O",
    "Pune H.O",
    "Pune New Bazar S.O",
    "Rupeenagar S.O",
    "Sachapir Street S.O",
    "Salisbury Park S.O",
    "Sangavi S.O",
    "Sasanenagar S.O",
    "Srpf S.O",
    "T.V. Nagar S.O",
    "Vadgaon Sheri S.O",
    "Vagholi S.O",
    "Vidyanagar S.O (Pune)",
    "Viman nagar S.O",
    "Vishrantwadi S.O",
    "Wanowarie S.O",
    "Yamunanagar S.O",
    "Yerwada S.O"
]

document.getElementById('file-input').addEventListener('change', handleFileSelect, false);
document.getElementById('sheet-drop-down').addEventListener('change', showCurrentSheet, false);
document.getElementById('merge-button').addEventListener('click', mergeSheets, false);
document.getElementById('reset').addEventListener('click', showCurrentSheet, false);
document.getElementById('office-drop-down').addEventListener('change', updateDataTable, false);
document.getElementById('sum').addEventListener('click', CalculateSum, false);
document.getElementById('sort').addEventListener('click', sortTable, false);
document.getElementById('save-as-excel').addEventListener('click', saveAsExcel, false);

let bSumClicked = false;
let pWorkbook = null;

function resetOfficeDropDown()
{
    const officeDropDown = document.getElementById('office-drop-down');
    if (officeDropDown.options.length > 0) {
        officeDropDown.selectedIndex = 0;
    }
}

function resetSortDropDown()
{
    const sortDropDown = document.getElementById('sort-drop-down');
    if (sortDropDown.options.length > 0) {
        sortDropDown.selectedIndex = 0;
    }
}

// Function to open selected file
function handleFileSelect(event)
{
    const file = event.target.files[0];
    if(!file){
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e)
    {
        const data = new Uint8Array(e.target.result);
        pWorkbook = XLSX.read(data, {type:'array'});

        const sheetSelect = document.getElementById('sheet-drop-down');
        const mergeButton = document.getElementById('merge-button');
        sheetSelect.innerHTML = "";
        pWorkbook.SheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = index; // Use index as the value
            option.textContent = sheetName; // Set the option text
            sheetSelect.appendChild(option);
        });

        sheetSelect.style.display = pWorkbook.SheetNames.length > 1 ? 'inline-block' : 'none';
        mergeButton.style.display = pWorkbook.SheetNames.length > 1 ? 'inline-block' : 'none';

        showCurrentSheet();

    };

    reader.readAsArrayBuffer(file);

    document.getElementById('reset').style.display = 'inline-block';
    document.getElementById('office-ui').style.display = 'inline-block';
    document.getElementById('sum').style.display = 'inline-block';
    document.getElementById('sorting-ui').style.display = 'inline-block';
    document.getElementById('save-as-excel').style.display = 'inline-block';
}

function showCurrentSheet(){
    const sheetDropdown = document.getElementById('sheet-drop-down');
    const pageIndex = sheetDropdown.value;
    const sheetName = pWorkbook.SheetNames[pageIndex];
    const worksheet = pWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, {header:1, defval: "" });

    displayData(jsonData);

    bSumClicked = false;
    resetOfficeDropDown();
    resetSortDropDown();
    sheetDropdown.disabled = false; 
}

// Function to display given data
function displayData(data)
{
    const tableBody = document.getElementById('data-table').getElementsByTagName('tbody')[0];
    const tableHead = document.getElementById('data-table').getElementsByTagName('thead')[0];

    tableHead.innerHTML = "";
    tableBody.innerHTML = "";

    let bTitleFound = false;
    data.forEach((row) => {
        const tr = document.createElement('tr');

        //skip the row which has < 4 values
        let iNbCellValue = 0;
        row.forEach((cell) => 
        {
            if(cell)
                iNbCellValue = iNbCellValue + 1;
        });

        if(iNbCellValue > 3)
        {
            row.forEach((cell) => {
                const td = document.createElement(bTitleFound ? 'td' : 'th');
                td.textContent = cell;
                tr.appendChild(td);
            });

             // If this is the first row, append to the thead
             if (!bTitleFound)
            {
                tableHead.appendChild(tr);
                bTitleFound = true;
            } 
            else
             {
                 tableBody.appendChild(tr);
             }              
        }
        });

    populateDropdown();
    copyDataToMasterTable();
}

// Function to calculate sums and update the table
 function CalculateSum()
 {
    const table = document.getElementById('data-table');
    const rows = table.getElementsByTagName('tr');

    const vecColHavingIntValues = findIntegerColumns(rows);

    if(!bSumClicked)
        createCheckboxes(vecColHavingIntValues);

    const totalHeader = rows[0].getElementsByTagName('th');
    for (let i = 0; i < rows.length; i++) {

        if(i === 0)
        {
            if (!bSumClicked && totalHeader.length > 0) 
            {
                const sumCell = document.createElement('th');
                sumCell.textContent = "Total";
                rows[i].appendChild(sumCell);
            }
        }
        else
        {
            const cells = rows[i].getElementsByTagName('td');
            if (cells.length > 0) 
            { 
                let sum = 0;
                vecColHavingIntValues.forEach(index => {
                const checkbox = document.getElementById(`checkbox-${index}`);
                if (checkbox && checkbox.checked)
                    {
                        const cellValue = parseFloat(cells[index]?.textContent || 0);
                        if (!isNaN(cellValue)) {
                            sum += cellValue;
                        }
                    }
                });
    
                if (bSumClicked && cells.length >= totalHeader.length) 
                {
                    cells[cells.length - 1].textContent = sum;
                }
                else
                {
                    const sumCell = document.createElement('td');
                    sumCell.textContent = sum;
                    rows[i].appendChild(sumCell);
                }
            }
        }
    }

    populateDropdown();
    bSumClicked = true;
};

// Function to check boxes in title row
function createCheckboxes(vecColHavingIntValues) {

    const table = document.getElementById('data-table');
    const rows = table.getElementsByTagName('tr');
    const headers = Array.from(rows[0].getElementsByTagName('th')).map(td => td.textContent.trim());

   const headerRow = table.getElementsByTagName('thead')[0].getElementsByTagName('tr')[0];
   headerRow.innerHTML = ""; // Clear existing headers

   headers.forEach((header, cellIndex) => {
      const th = document.createElement('th');
      if(vecColHavingIntValues.includes(cellIndex))
      {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `checkbox-${cellIndex}`;
        checkbox.checked = true; // Check all boxes by default
        checkbox.addEventListener('change', checkBoxClicked);
        const label = document.createElement('label');
        label.htmlFor = checkbox.id;
        label.textContent = header; // Set the checkbox label
        th.appendChild(checkbox);
        th.appendChild(label);
      }
      else
      {
        th.textContent = header;
      }

      headerRow.appendChild(th); // Append header cell with checkbox
   });
}

// Check if the cellValue is a number (including NaN check)
function isCellValueValid(cellValue) {
    if (cellValue === "" || cellValue === null) {
        return false; // Empty or null values are not valid
    }
    return !isNaN(cellValue) || typeof cellValue === 'string'; // Valid if it's a number or a non-empty string
}

// Finds the Columns have int values
function findIntegerColumns(tabel) 
{
    const numRow = tabel.length;
    const integerColumns = [];

    const titleCell = tabel[0].getElementsByTagName('th');
    let iNbTitleCell = bSumClicked ? titleCell.length - 1 : titleCell.length;

    for(let iCell = 0; iCell < iNbTitleCell; iCell++)
    {
        let isIntegerColumn = true;
        for (let iRow = 1; iRow < numRow ; iRow++) 
            {
                const cell = tabel[iRow].getElementsByTagName('td');

                if(iCell >= cell.length)
                {
                    isIntegerColumn = false;
                    break;
                }
        
                const cellValue = parseFloat(cell[iCell]?.textContent || 0);
                if (!Number.isInteger(cellValue))
                {
                    isIntegerColumn = false;
                    break;
                }

            }

            // If the column has only integers, add the index to the result
            if (isIntegerColumn) 
            {
                integerColumns.push(iCell);
            }
    }

    return integerColumns;
}

// Update the drop down
function populateDropdown() {
    const table = document.getElementById('data-table');
    const headerSelect = document.getElementById('sort-drop-down');
    headerSelect.innerHTML = ""; // Clear existing options

    const headers = Array.from(table.getElementsByTagName('th'));
    headers.forEach((th, index) => {
        const option = document.createElement('option');
        option.value = index; // Use the index as the value
        option.textContent = th.textContent; // Set the option text
        headerSelect.appendChild(option);
    });
}

// Sort the given table based on the value of cellIndex
function sortTable() {
    const table = document.getElementById('data-table');
    const headerSelect = document.getElementById('sort-drop-down');
    const cellIndex = parseInt(headerSelect.value, 10);
    const sortOrder = document.querySelector('input[name="sortOrder"]:checked').value;

    const rows = Array.from(table.getElementsByTagName('tr')).slice(1); // Skip header row
    if(rows.length < 2)
        return;

    // Function to determine the type of value and return a comparable value
    const getComparableValue = (cell) => {
        const value = cell.textContent.trim();

        // Try to parse as a number
        const numberValue = parseFloat(value);
        if (!isNaN(numberValue)) {
            return { value: numberValue, type: 'number' }; // Return number and type
        }

        // Return as string if it's not a number
        return { value: value.toLowerCase(), type: 'string' }; // Use lower case for string comparison

    };
        
    // Sort rows based on the specified cell index
    rows.sort((a, b) => {
        const aValue = getComparableValue(a.getElementsByTagName('td')[cellIndex]);
        const bValue = getComparableValue(b.getElementsByTagName('td')[cellIndex]);

        // Check if either row contains the value "total"
        const aContainsTotal = a.textContent.toLowerCase().includes("total");
        const bContainsTotal = b.textContent.toLowerCase().includes("total");
    
        // If either row contains "total", move it to the end
        if (aContainsTotal && !bContainsTotal) {
            return 1; // a goes after b
        }
        if (!aContainsTotal && bContainsTotal) {
            return -1; // a goes before b
        }

        let comparisonResult;
        if (aValue.type === 'number' && bValue.type === 'number') {
            comparisonResult = aValue.value - bValue.value; // Numeric comparison
        } else if (aValue.type === 'string' && bValue.type === 'string') {
            comparisonResult = aValue.value.localeCompare(bValue.value); // String comparison
        } else if (aValue.type === 'number') {
            comparisonResult = -1; // Numbers come before strings
        } else {
            comparisonResult = 1; // Strings come after numbers
        }

        // Adjust based on sort order
        return sortOrder === 'ascending' ? comparisonResult : -comparisonResult;
    });

    // Remove existing rows from the table body
    const tableBody = table.getElementsByTagName('tbody')[0];
    tableBody.innerHTML = "";

    // Append sorted rows back to the table body
    rows.forEach(row => {
        tableBody.appendChild(row);
    });
}

// Function to save the table data as an Excel file
function saveAsExcel() {
    const fileName = prompt("Enter the filename (without extension):", "table_data");
    if (!fileName) return; // Exit if no filename is provided

    const table = document.getElementById('data-table');
    const rows = Array.from(table.getElementsByTagName('tr'));
    
    // Convert table rows to a 2D array
    const data = rows.map(row => {
        return Array.from(row.getElementsByTagName('th')).map(th => th.textContent) // For header
            .concat(Array.from(row.getElementsByTagName('td')).map(td => td.textContent)); // For data
    }).filter(row => row.length > 0); // Filter out empty rows

    const worksheet = XLSX.utils.aoa_to_sheet(data); // Create a worksheet from the 2D array
    const workbook = XLSX.utils.book_new(); // Create a new workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1"); // Append the worksheet to the workbook

    // Export the workbook to a file
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
}

function updateDataTable() {
    const masterTable = document.getElementById('master-table');
    const dataTableBody = document.getElementById('data-table').getElementsByTagName('tbody')[0];
    const selectedOption = document.getElementById('office-drop-down').value;
    
    const numberOfColumns = document.getElementById('data-table').getElementsByTagName('th').length;

    const rowsToClear = dataTableBody.rows.length;
    for (let i = rowsToClear - 1; i >= 0; i--) { // Start from the last row and skip the first (header) row
        dataTableBody.deleteRow(i);
    }

    // Get rows from the main data table
    const rows = Array.from(masterTable.getElementsByTagName('tr')).slice(1);

    if(selectedOption === 'show-main-table')
    {
        rows.forEach(row => {
            const newRow = row.cloneNode(true);
            dataTableBody.appendChild(newRow);
        });

    }
    else if(selectedOption === 'show-so-only')
    {
        rows.forEach(row => {
            const firstCellValue = row.getElementsByTagName('td')[0]?.textContent.trim();
            if (vecSO.includes(firstCellValue)) {
                const newRow = row.cloneNode(true);
                dataTableBody.appendChild(newRow);
            }
        });
    }
    else if(selectedOption === 'show-bo-only')
    {
        rows.forEach(row => {
            const firstCellValue = row.getElementsByTagName('td')[0]?.textContent.trim();
            if (vecBO.includes(firstCellValue)) {
                const newRow = row.cloneNode(true); // Clone the row to keep its content
                dataTableBody.appendChild(newRow);
            }
        });
    }
    else if(selectedOption === 'show-all-so' || selectedOption === 'show-all-bo')
    {
        let vecOffice = (selectedOption === 'show-all-so') ? [...vecSO] : [...vecBO];
        for (const office of vecOffice) {
            const foundRow = rows.find(row => row.getElementsByTagName('td')[0]?.textContent.toLowerCase() === office.toLowerCase());

            if (foundRow) {
                // If found, copy the row
                const newRow = dataTableBody.insertRow();
                Array.from(foundRow.cells).forEach(cell => {
                    const newCell = newRow.insertCell();
                    newCell.textContent = cell.textContent; // Copy cell data to the new table
                });
            } else {
                // If not found, create a new row with the S.O. name
                const newRow = dataTableBody.insertRow();
                const firstCell = newRow.insertCell();
                firstCell.textContent = office; // Set S.O. name
                // Add empty cells for the rest
                for (let i = 1; i < numberOfColumns; i++) { // Adjust numberOfColumns as per your table
                    newRow.insertCell(); 
                }
            }
        }
    }
    else
    {
        rows.forEach(row => {
            const newRow = row.cloneNode(true);
            dataTableBody.appendChild(newRow);
        });

        let vecOffice = (selectedOption === 'add-all-so') ? [...vecSO] : [...vecBO];
        if(selectedOption === 'add-all-so-bo')
        {
            vecOffice.push(...vecSO);
        }

        for (const office of vecOffice) {
            const foundRow = rows.find(row => row.getElementsByTagName('td')[0]?.textContent.toLowerCase() === office.toLowerCase());

            if (!foundRow) {
                // If not found, create a new row with the S.O. name
                const newRow = dataTableBody.insertRow();
                const firstCell = newRow.insertCell();
                firstCell.textContent = office; // Set S.O. name
                // Add empty cells for the rest
                for (let i = 1; i < numberOfColumns; i++) { // Adjust numberOfColumns as per your table
                    newRow.insertCell(); 
                }
            }
        }
    }

    const sortDropDown = document.getElementById('sort-drop-down');
    if (sortDropDown.options.length > 0) {
        sortDropDown.selectedIndex = 0;
    }

    sortTable();
}

function copyDataToMasterTable(){

    const dataTable = document.getElementById('data-table');
    const masterTableBody = document.getElementById('master-table').getElementsByTagName('tbody')[0];
    masterTableBody.innerHTML = "";

    // Get rows from the main data table
    const rows = Array.from(dataTable.getElementsByTagName('tr')).slice(1); // Skip header row

    rows.forEach(row => {
        const newRow = row.cloneNode(true);
        masterTableBody.appendChild(newRow);
    });
}

function checkBoxClicked() {

    if(!bSumClicked)
        return;

    const table = document.getElementById('data-table');
    const rows = table.getElementsByTagName('tr');

    // Loop through all rows except the header (first row)
    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName('td');

        // Only act if there are cells in the row
        if (cells.length > 0) {
            const lastCell = cells[cells.length - 1];
            if (lastCell) {
                lastCell.textContent = ""; // Clear the last cell's content
            }
        }
    }
}

function mergeSheets() {
    const sheetDropdown = document.getElementById('sheet-drop-down');
    sheetDropdown.disabled = true; 

    const sheetNames = pWorkbook.SheetNames;
    if (sheetNames.length <= 1) {
        return; // Do nothing if only one sheet
    }

    const combinedData = [];
    const firstSheetColumnCount = getNumberOfColumns(0);
    for (let i = 0; i < sheetNames.length; i++) {
        const currentSheetColumnCount = getNumberOfColumns(i);

        // If a sheet has a different number of columns, show a warning
        if (currentSheetColumnCount !== firstSheetColumnCount) {
            alert("Warning: Not all sheets have the same number of columns.");
            return; // Exit the function
        }

        // Get the data from the current sheet
        const worksheet = pWorkbook.Sheets[sheetNames[i]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header:1, defval: "" });

        // Append data to the combinedData array
        combinedData.push(...jsonData);
    }

    displayCombinedData(combinedData);
}

function getNumberOfColumns(sheetIndex) {
    const sheetName = pWorkbook.SheetNames[sheetIndex];
    const worksheet = pWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    let maxColumns = 0;
    jsonData.forEach(row => {
        if (row.length > maxColumns) {
            maxColumns = row.length; // Update the maximum columns found
        }
    });

    return maxColumns; // Return the maximum number of columns found
}

// Function to display combined data in the data table
function displayCombinedData(data) {
    const tableBody = document.getElementById('data-table').getElementsByTagName('tbody')[0];
    tableBody.innerHTML = ""; // Clear existing rows

    // Iterate over the combined data and create rows
    data.forEach(row => {
        if (row.length >= 4)
        {
            const firstCellValue = row[0];
            const totalPattern = /total/i; // Regular expression for "total", case insensitive
            const namePattern = /name/i; // Regular expression for "total", case insensitive
            if (!firstCellValue || totalPattern.test(firstCellValue)  || namePattern.test(firstCellValue)) {
                return; // Skip this row if any condition is met
            }

            const existingRow = Array.from(tableBody.getElementsByTagName('tr')).find(existingRow => 
                existingRow.cells[0].textContent === firstCellValue);

            if (existingRow)
            {
                for (let i = 1; i < row.length; i++) {
                    const existingCellValue = existingRow.cells[i]?.textContent.trim();
                    const newCellValue = typeof row[i] === 'string' ? row[i].trim() : row[i];

                    const existingNum = parseFloat(existingCellValue) || null;
                    const newNum = parseFloat(newCellValue) || null;

                    if (existingNum !== null && newNum !== null) {
                        // 1) If both are integers, add them
                        existingRow.cells[i].textContent = existingNum + newNum;
                    } else if (existingCellValue === "" || newCellValue === "") {
                        // 2) If any one of them is empty, take the non-empty value
                        existingRow.cells[i].textContent = existingCellValue || newCellValue;
                    } else {
                        // 3) If both are strings, keep the first value
                        existingRow.cells[i].textContent = existingCellValue; // Keep the existing value
                    }
                }
            }
            else{
                const tr = document.createElement('tr');
                row.forEach(cell => {
                    const td = document.createElement('td');
                    td.textContent = cell; // Set cell value
                    tr.appendChild(td);
                });
                tableBody.appendChild(tr); // Append the new row to the table body
            }


        }

    });

    // Optional: Call any function to sort or format the table after populating it
    sortTable(); // Call your sorting function if needed
}
