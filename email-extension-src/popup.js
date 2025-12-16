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
    let foundRowIndex = -1;
    let currentUpdateRange = '';

    // 1. Initialize: Get Auth Token and Load Sheets
    showStatus('Initializing...', 'normal');
    // Using explicit interactive: true might fail if triggered on load without user gesture in some browser versions.
    // Try interactive: false first, then if needed prompt user?
    // Actually, for a popup, interactive: true usually prompts the login window if needed.
    getAuthToken(true, (token) => {
        if (token) {
            loadSheets(token);
        } else {
            // It's possible the user closed the auth window or it failed.
            showStatus('Authorization Check Failed. Check console for details.', 'error');
            // Adding a manual retry button could be helpful here if we were fancy, 
            // but for now let's just show the error.
        }
    });

    // 2. Search Handler
    searchBtn.addEventListener('click', () => {
        const email = emailInput.value.trim();
        const sheetName = sheetSelect.value; // This might be empty if sheets failed to load

        if (!email) {
            showStatus('Please enter an email.', 'error');
            return;
        }
        if (!sheetName) {
            // If sheets haven't loaded, user can't select one.
            showStatus('Please select a sheet (Wait for sheets to load).', 'error');
            return;
        }

        getAuthToken(false, (token) => {
            if (token) {
                searchForEmail(token, sheetName, email);
            } else {
                showStatus('Auth token missing during search.', 'error');
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
            } else {
                showStatus('Auth token missing during update.', 'error');
            }
        });
    });

    // --- Helper Functions ---

    function getAuthToken(interactive, callback) {
        try {
            chrome.identity.getAuthToken({ interactive: interactive }, (token) => {
                if (chrome.runtime.lastError) {
                    console.error('getAuthToken Error:', chrome.runtime.lastError);
                    showStatus(`Auth Error: ${chrome.runtime.lastError.message}`, 'error');
                    callback(null);
                } else {
                    console.log('Token received');
                    callback(token);
                }
            });
        } catch (e) {
            console.error('Exception in getAuthToken:', e);
            showStatus(`Auth Exception: ${e.message}`, 'error');
            callback(null);
        }
    }

    function showStatus(msg, type = 'normal') {
        // Also log to console so user can inspect
        if (type === 'error') console.error(msg);
        else console.log(msg);

        statusMessage.textContent = msg;
        statusMessage.className = 'status ' + type;
    }

    async function loadSheets(token) {
        showStatus('Fetching sheets...');
        try {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?fields=sheets.properties.title`;
            console.log('Fetching:', url);

            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });

            console.log('Response status:', response.status);

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API Error ${response.status}: ${errorText}`);
            }

            const data = await response.json();

            if (data.error) {
                throw new Error(data.error.message);
            }

            console.log('Sheets data received:', data);

            // Clear existing options
            sheetSelect.innerHTML = '<option value="" disabled selected>Select a sheet</option>';

            if (data.sheets && data.sheets.length > 0) {
                data.sheets.forEach(sheet => {
                    const title = sheet.properties.title;
                    const option = document.createElement('option');
                    option.value = title;
                    option.textContent = title;
                    sheetSelect.appendChild(option);
                });
                showStatus(''); // Clear 'Loading...' message on success
            } else {
                showStatus('No sheets found in this spreadsheet.', 'error');
            }

        } catch (err) {
            showStatus(`Load Sheets Failed: ${err.message}`, 'error');
        }
    }

    async function searchForEmail(token, sheetName, email) {
        showStatus(`Searching in "${sheetName}"...`);
        resultSection.classList.add('hidden');
        currentSheetName = sheetName;

        try {
            const range = `${sheetName}!A:Z`;
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}`;

            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API Error ${response.status}: ${errorText}`);
            }

            const data = await response.json();

            if (data.error) throw new Error(data.error.message);
            if (!data.values || data.values.length === 0) {
                showStatus('Sheet is empty.', 'error');
                return;
            }

            const headers = data.values[0];
            // Flexible header matching
            const emailIndex = headers.findIndex(h => h && h.trim().toLowerCase() === 'email');
            const respondedIndex = headers.findIndex(h => h && h.trim().toLowerCase() === 'responded?');
            const nameIndex = headers.findIndex(h => h && (h.trim().toLowerCase() === 'first name' || h.trim().toLowerCase() === 'name'));

            if (emailIndex === -1) {
                showStatus('Column "Email" not found in this sheet.', 'error');
                return;
            }
            if (respondedIndex === -1) {
                showStatus('Column "Responded?" not found in this sheet.', 'error');
                return;
            }

            let matchRow = -1;
            let matchData = null;
            const searchEmail = email.toLowerCase().trim();

            for (let i = 1; i < data.values.length; i++) {
                const row = data.values[i];
                // Ensure row has enough columns
                const rowEmail = (row[emailIndex]) ? String(row[emailIndex]).toLowerCase().trim() : '';
                if (rowEmail === searchEmail) {
                    matchRow = i;
                    matchData = row;
                    break;
                }
            }

            if (matchRow !== -1) {
                foundRowIndex = matchRow + 1;
                const name = (nameIndex !== -1 && matchData[nameIndex]) ? matchData[nameIndex] : '(Name not found)';

                contactNameSpan.textContent = name;
                contactEmailSpan.textContent = matchData[emailIndex];

                const currentStatus = (matchData[respondedIndex]) ? String(matchData[respondedIndex]).toLowerCase() : '';

                let optionFound = false;
                for (let opts of responseSelect.options) {
                    if (opts.value === currentStatus) {
                        responseSelect.value = currentStatus;
                        optionFound = true;
                    }
                }
                if (!optionFound) {
                    responseSelect.value = "";
                }

                const respondedColLetter = getColumnLetter(respondedIndex);
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

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`API Error ${response.status}: ${errorText}`);
            }

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
