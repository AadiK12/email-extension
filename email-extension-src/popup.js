const SPREADSHEET_ID = '1UsOeACSNf5e3DvYItDjzc6pgtf5BaPaB4yXXACP7sl8'; // Contacts Master Sheet ID

document.addEventListener('DOMContentLoaded', () => {
    const sheetSelect = document.getElementById('sheetSelect');
    const emailInput = document.getElementById('emailInput');
    const searchBtn = document.getElementById('searchBtn');
    const statusMessage = document.getElementById('statusMessage');
    const resultSection = document.getElementById('resultSection');
    const contactNameSpan = document.getElementById('contactName');
    const contactEmailSpan = document.getElementById('contactEmail');
    const responseSelect = document.getElementById('responseSelect');
    const updateBtn = document.getElementById('updateBtn');

    let currentSheetName = '';
    let foundRowIndex = -1; // 0-indexed relative to sheet data? API uses 0-indexed or A1 notation? We'll use A1.

    // 1. Initialize: Get Auth Token and Load Sheets
    getAuthToken(true, (token) => {
        if (token) {
            loadSheets(token);
        } else {
            showStatus('Please click extension to authorize.', 'error');
        }
    });

    // 2. Search Handler
    searchBtn.addEventListener('click', () => {
        const email = emailInput.value.trim();
        const sheetName = sheetSelect.value;

        if (!email || !sheetName) {
            showStatus('Please enter an email and select a sheet.', 'error');
            return;
        }

        getAuthToken(false, (token) => {
            if (token) {
                searchForEmail(token, sheetName, email);
            }
        });
    });

    // 3. Update Handler
    updateBtn.addEventListener('click', () => {
        const newStatus = responseSelect.value;
        if (!newStatus) {
            showStatus('Please select a status.', 'error');
            return;
        }

        getAuthToken(false, (token) => {
            if (token) {
                updateResponseStatus(token, currentSheetName, foundRowIndex, newStatus);
            }
        });
    });

    // --- Helper Functions ---

    function getAuthToken(interactive, callback) {
        chrome.identity.getAuthToken({ interactive: interactive }, (token) => {
            if (chrome.runtime.lastError) {
                console.error(chrome.runtime.lastError);
                showStatus('Authentication failed. See console.', 'error');
                callback(null);
            } else {
                callback(token);
            }
        });
    }

    function showStatus(msg, type = 'normal') {
        statusMessage.textContent = msg;
        statusMessage.className = 'status ' + type;
    }

    async function loadSheets(token) {
        showStatus('Loading sheets...');
        try {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?fields=sheets.properties.title`;
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            const data = await response.json();

            if (data.error) {
                throw new Error(data.error.message);
            }

            sheetSelect.innerHTML = '<option value="" disabled selected>Select a sheet</option>';
            data.sheets.forEach(sheet => {
                const title = sheet.properties.title;
                const option = document.createElement('option');
                option.value = title;
                option.textContent = title;
                sheetSelect.appendChild(option);
            });
            showStatus('');
        } catch (err) {
            showStatus(`Error loading sheets: ${err.message}`, 'error');
        }
    }

    async function searchForEmail(token, sheetName, email) {
        showStatus(`Searching in "${sheetName}"...`);
        resultSection.classList.add('hidden');
        currentSheetName = sheetName;
        
        try {
            // Read the entire sheet's values (or substantial chunk)
            // Ideally we'd read just headers first to find columns, then data.
            // For simplicity/speed, we'll read A1:Z1 (headers) then A:Z.
            
            const range = `${sheetName}!A:Z`; // Assumption: Data is within cols A-Z.
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}`;
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            const data = await response.json();

            if (data.error) throw new Error(data.error.message);
            if (!data.values || data.values.length === 0) {
                showStatus('Sheet is empty.', 'error');
                return;
            }

            const headers = data.values[0];
            const emailIndex = headers.findIndex(h => h.trim() === 'Email');
            const respondedIndex = headers.findIndex(h => h.trim() === 'Responded?');
            const nameIndex = headers.findIndex(h => h.trim() === 'First Name' || h.trim() === 'Name'); // Try to match common name headers

            if (emailIndex === -1) {
                showStatus('Column "Email" not found in this sheet.', 'error');
                return;
            }
            if (respondedIndex === -1) {
                showStatus('Column "Responded?" not found in this sheet.', 'error');
                return;
            }

            // Client-side search for the email
            let matchRow = -1;
            let matchData = null;
            const searchEmail = email.toLowerCase().trim();

            for (let i = 1; i < data.values.length; i++) {
                const row = data.values[i];
                const rowEmail = row[emailIndex] ? row[emailIndex].toLowerCase().trim() : '';
                if (rowEmail === searchEmail) {
                    matchRow = i; // 0-based index in the values array
                    matchData = row;
                    break;
                }
            }

            if (matchRow !== -1) {
                foundRowIndex = matchRow + 1; // 1-based row number for Sheets API A1 notation
                const name = (nameIndex !== -1 && matchData[nameIndex]) ? matchData[nameIndex] : '(Name not found)';
                
                contactNameSpan.textContent = name;
                contactEmailSpan.textContent = matchData[emailIndex];
                
                // Set current status if available
                const currentStatus = matchData[respondedIndex] ? matchData[respondedIndex].toLowerCase() : '';
                // Try to set the select value, or default to empty
                let optionFound = false;
                for(let opts of responseSelect.options){
                     if(opts.value === currentStatus) {
                         responseSelect.value = currentStatus;
                         optionFound = true;
                     }
                }
                if(!optionFound && currentStatus) {
                     // If existing status isn't in our list, maybe add it or just leave select blank?
                     // Leaving blank/default is safer than overwriting visually.
                     responseSelect.value = "";
                }

                // Store column letter for update
                // 0 -> A, 1 -> B ... 
                const respondedColLetter = getColumnLetter(respondedIndex);
                // Save specific cell range: e.g. Sheet1!C5
                currentUpdateRange = `${sheetName}!${respondedColLetter}${foundRowIndex}`;

                resultSection.classList.remove('hidden');
                showStatus('Contact found.', 'success');
            } else {
                showStatus('Email not found in this sheet.', 'error');
            }

        } catch (err) {
            showStatus(`Search error: ${err.message}`, 'error');
        }
    }

    let currentUpdateRange = '';

    async function updateResponseStatus(token, sheetName, rowIndex, newStatus) {
        if (!currentUpdateRange) return;

        showStatus('Updating...');
        try {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(currentUpdateRange)}?valueInputOption=USER_ENTERED`;
            
            const body = {
                range: currentUpdateRange,
                majorDimension: 'ROWS',
                values: [[newStatus]]
            };

            const response = await fetch(url, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(body)
            });

            const data = await response.json();

            if (data.error) throw new Error(data.error.message);

            showStatus('Status updated successfully!', 'success');

        } catch (err) {
            showStatus(`Update error: ${err.message}`, 'error');
        }
    }

    function getColumnLetter(index) {
        let temp, letter = '';
        while (index >= 0) {
            temp = (index) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            index = Math.floor((index) / 26) - 1;
        }
        return letter;
    }
});
