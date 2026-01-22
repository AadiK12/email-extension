const SPREADSHEETS = [
    { id: '1UsOeACSNf5e3DvYItDjzc6pgtf5BaPaB4yXXACP7sl8', name: 'Master 1' },
    { id: '1XkXb7QWgZzpKddQdyvxK7NrfsuSJB5ZARWnRroHWav0', name: 'Master 2' }
];

const SALES_MASTER_ID = '16yU-doSAcsxTOvNIbkw2RmFa059P6MCDPLmPBOMZuaQ';
const SENT_LOG_SHEET_NAME = 'SentLog';

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

    // Calendly Elements
    const calendlyTokenInput = document.getElementById('calendlyToken');
    const saveTokenBtn = document.getElementById('saveTokenBtn');
    const toggleTokenBtn = document.getElementById('toggleTokenBtn');
    const tokenInputSection = document.getElementById('tokenInputSection');
    const calendlySection = document.getElementById('calendlySection');
    const eventTypeSelect = document.getElementById('eventTypeSelect');
    const calNameInput = document.getElementById('calNameInput');
    const calEmailInput = document.getElementById('calEmailInput');
    const generateLinkBtn = document.getElementById('generateLinkBtn');
    const calendlyResult = document.getElementById('calendlyResult');
    const generatedLinkDisplay = document.getElementById('generatedLinkDisplay');
    const copyLinkBtn = document.getElementById('copyLinkBtn');

    let calendlyUserUri = null;

    let currentSpreadsheetId = '';
    let currentSheetName = '';
    let foundRowIndex = -1;
    let currentUpdateRange = '';
    let allSheetsContext = []; // Stores { spreadsheetId, sheetName } for all available sheets

    // 1. Initialize: Get Auth Token and Load Sheets
    showStatus('Initializing...', 'normal');

    // Check for Calendly Token
    chrome.storage.local.get(['calendlyToken'], (result) => {
        if (result.calendlyToken) {
            calendlyTokenInput.value = result.calendlyToken;
            initCalendly(result.calendlyToken);
        } else {
            tokenInputSection.classList.remove('hidden');
        }
    });
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
        const selection = sheetSelect.value; // This might be empty if sheets failed to load

        if (!email) {
            showStatus('Please enter an email.', 'error');
            return;
        }
        if (!selection) {
            showStatus('Please select a sheet (Wait for sheets to load).', 'error');
            return;
        }

        if (selection === 'ALL') {
            getAuthToken(false, (token) => {
                if (token) {
                    // New logic: Search via SentLog first
                    searchViaSentLog(token, email);
                } else {
                    showStatus('Auth token missing during search.', 'error');
                }
            });
            return;
        }

        const [spreadsheetId, sheetName] = selection.split('|||'); // Existing single sheet logic

        getAuthToken(false, (token) => {
            if (token) {
                searchForEmail(token, spreadsheetId, sheetName, email);
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
                updateResponseStatus(token, currentSpreadsheetId, currentSheetName, foundRowIndex, newStatus);
            } else {
                showStatus('Auth token missing during update.', 'error');
            }
        });
    });

    // 4. Calendly Handlers
    toggleTokenBtn.addEventListener('click', () => {
        tokenInputSection.classList.toggle('hidden');
    });

    saveTokenBtn.addEventListener('click', () => {
        const token = calendlyTokenInput.value.trim();
        if (token) {
            chrome.storage.local.set({ calendlyToken: token }, () => {
                showStatus('Calendly token saved.', 'success');
                tokenInputSection.classList.add('hidden');
                initCalendly(token);
            });
        }
    });

    generateLinkBtn.addEventListener('click', () => {
        const token = calendlyTokenInput.value.trim();
        const eventTypeUri = eventTypeSelect.value;
        const name = calNameInput.value.trim();
        const email = calEmailInput.value.trim();

        if (!token) return showStatus('Calendly token missing.', 'error');
        if (!eventTypeUri) return showStatus('Please select an event type.', 'error');

        createCalendlyLink(token, eventTypeUri, name, email);
    });

    copyLinkBtn.addEventListener('click', () => {
        generatedLinkDisplay.select();
        document.execCommand('copy');
        showStatus('Link copied!', 'success');
    });

    // 5. Template Handlers
    const templateSelect = document.getElementById('templateSelect');
    const newTemplateBtn = document.getElementById('newTemplateBtn');
    const templateEditor = document.getElementById('templateEditor');
    const templateTitleObj = document.getElementById('templateTitle');
    const templateContentObj = document.getElementById('templateContent');
    const saveTemplateBtn = document.getElementById('saveTemplateBtn');
    const copyTemplateBtn = document.getElementById('copyTemplateBtn');
    const deleteTemplateBtn = document.getElementById('deleteTemplateBtn');

    let templates = [];
    let currentTemplateId = null;

    // Load templates on startup
    loadTemplates();

    newTemplateBtn.addEventListener('click', () => {
        currentTemplateId = null;
        templateTitleObj.value = '';
        templateContentObj.value = '';
        templateSelect.value = '';
        templateEditor.classList.remove('hidden');
        templateTitleObj.focus();
    });

    templateSelect.addEventListener('change', () => {
        const id = templateSelect.value;
        if (id) {
            const tmpl = templates.find(t => t.id === id);
            if (tmpl) {
                currentTemplateId = tmpl.id;
                templateTitleObj.value = tmpl.title;
                templateContentObj.value = tmpl.content;
                templateEditor.classList.remove('hidden');
            }
        } else {
            templateEditor.classList.add('hidden');
        }
    });

    saveTemplateBtn.addEventListener('click', () => {
        const title = templateTitleObj.value.trim();
        const content = templateContentObj.value.trim();

        if (!title) {
            showStatus('Please enter a template title.', 'error');
            return;
        }

        saveTemplate(title, content);
    });

    deleteTemplateBtn.addEventListener('click', () => {
        if (currentTemplateId) {
            if (confirm('Are you sure you want to delete this template?')) {
                deleteTemplate(currentTemplateId);
            }
        }
    });

    copyTemplateBtn.addEventListener('click', () => {
        if (!templateContentObj.value) return;
        templateContentObj.select();
        document.execCommand('copy');
        showStatus('Template copied to clipboard!', 'success');
    });

    function loadTemplates() {
        chrome.storage.local.get(['emailTemplates'], (result) => {
            templates = result.emailTemplates || [];
            renderTemplateList();
        });
    }

    function renderTemplateList() {
        templateSelect.innerHTML = '<option value="">Select a template...</option>';
        templates.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t.id;
            opt.textContent = t.title;
            templateSelect.appendChild(opt);
        });

        if (currentTemplateId) {
            templateSelect.value = currentTemplateId;
        }
    }

    function saveTemplate(title, content) {
        if (currentTemplateId) {
            // Update existing
            const index = templates.findIndex(t => t.id === currentTemplateId);
            if (index !== -1) {
                templates[index] = { id: currentTemplateId, title, content };
            }
        } else {
            // Create new
            const newId = Date.now().toString();
            templates.push({ id: newId, title, content });
            currentTemplateId = newId;
        }

        chrome.storage.local.set({ emailTemplates: templates }, () => {
            showStatus('Template saved.', 'success');
            renderTemplateList();
        });
    }

    function deleteTemplate(id) {
        templates = templates.filter(t => t.id !== id);
        chrome.storage.local.set({ emailTemplates: templates }, () => {
            showStatus('Template deleted.', 'success');
            currentTemplateId = null;
            templateTitleObj.value = '';
            templateContentObj.value = '';
            templateEditor.classList.add('hidden');
            renderTemplateList();
        });
    }

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
        sheetSelect.innerHTML = '<option value="" disabled>Select a sheet</option>';
        allSheetsContext = [];

        try {
            let totalSheetsFound = 0;

            for (const sheetObj of SPREADSHEETS) {
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetObj.id}?fields=sheets.properties.title`;
                console.log(`Fetching from ${sheetObj.name}:`, url);

                const response = await fetch(url, {
                    headers: { 'Authorization': `Bearer ${token}` }
                });

                if (response.ok) {
                    const data = await response.json();
                    if (data.sheets && data.sheets.length > 0) {
                        const optGroup = document.createElement('optgroup');
                        optGroup.label = sheetObj.name;

                        data.sheets.forEach(sheet => {
                            const title = sheet.properties.title;
                            const option = document.createElement('option');
                            option.value = `${sheetObj.id}|||${title}`;
                            option.textContent = title;
                            optGroup.appendChild(option);

                            // Store for global search
                            allSheetsContext.push({
                                spreadsheetId: sheetObj.id,
                                sheetName: title,
                                sourceName: sheetObj.name
                            });
                        });
                        sheetSelect.appendChild(optGroup);
                        totalSheetsFound += data.sheets.length;
                    }
                } else {
                    console.error(`Failed to load ${sheetObj.name}: ${response.status}`);
                }
            }

            if (totalSheetsFound > 0) {
                // Add "Search All" option at the top
                const allOption = document.createElement('option');
                allOption.value = "ALL";
                allOption.textContent = "Search All Sheets";
                allOption.selected = true;
                sheetSelect.insertBefore(allOption, sheetSelect.firstChild);

                showStatus('');
            } else {
                showStatus('No sheets found.', 'error');
            }

        } catch (err) {
            showStatus(`Load Sheets Failed: ${err.message}`, 'error');
        }
    }

    async function searchViaSentLog(token, email) {
        showStatus('Checking SentLog...', 'normal');
        resultSection.classList.add('hidden');

        try {
            // 1. Get Sheet Properties (to find total row count) and Headers
            const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SALES_MASTER_ID}?fields=sheets.properties`;

            const [metaRes, headerRes] = await Promise.all([
                fetch(metaUrl, { headers: { 'Authorization': `Bearer ${token}` } }),
                fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SALES_MASTER_ID}/values/${SENT_LOG_SHEET_NAME}!1:1`, {
                    headers: { 'Authorization': `Bearer ${token}` }
                })
            ]);

            if (!metaRes.ok) throw new Error('Failed to fetch Sales Master metadata');
            if (!headerRes.ok) throw new Error('Failed to fetch SentLog headers');

            const metaData = await metaRes.json();
            const headerData = await headerRes.json();

            // Find SentLog sheet properties
            const sentLogSheet = metaData.sheets.find(s => s.properties.title === SENT_LOG_SHEET_NAME);
            if (!sentLogSheet) throw new Error(`Sheet "${SENT_LOG_SHEET_NAME}" not found in Sales Master.`);

            const totalRows = sentLogSheet.properties.gridProperties.rowCount;

            // Determine Headers
            if (!headerData.values || headerData.values.length === 0) throw new Error('SentLog headers empty.');
            const headers = headerData.values[0];
            const emailIndex = headers.findIndex(h => h && h.trim().toLowerCase() === 'email');
            const sourceSheetIndex = headers.findIndex(h => h && h.trim().toLowerCase() === 'source sheet');

            if (emailIndex === -1) throw new Error('No "Email" column in SentLog.');
            if (sourceSheetIndex === -1) throw new Error('No "Source Sheet" column in SentLog.');

            // 2. Fetch last 1500 rows
            const rowsToFetch = 1500;
            // gridProperties.rowCount is the total dimensions. 
            // If the sheet is huge but empty at bottom, we might scan empty rows.
            // But relying on that is the only way without iterative checking.
            const startRow = Math.max(2, totalRows - rowsToFetch + 1);
            const range = `${SENT_LOG_SHEET_NAME}!A${startRow}:Z${totalRows}`;

            showStatus(`Scanning last ${rowsToFetch} rows of SentLog...`);

            const dataRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SALES_MASTER_ID}/values/${encodeURIComponent(range)}`, {
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (!dataRes.ok) throw new Error('Failed to fetch SentLog data.');
            const data = await dataRes.json();

            if (!data.values || data.values.length === 0) {
                showStatus('No data found in the last 1500 rows.', 'error');
                return;
            }

            // 3. Search Data (Iterate Backwards)
            const searchEmail = email.toLowerCase().trim();
            // Data values index 0 corresponds to startRow. 
            // We want to loop from the end (most recent).
            let foundSourceSheet = null;

            for (let i = data.values.length - 1; i >= 0; i--) {
                const row = data.values[i];
                const rowEmail = (row[emailIndex]) ? String(row[emailIndex]).toLowerCase().trim() : '';

                if (rowEmail === searchEmail) {
                    foundSourceSheet = (row[sourceSheetIndex]) ? String(row[sourceSheetIndex]).trim() : '';
                    break;
                }
            }

            if (foundSourceSheet) {
                showStatus(`Found in SentLog! Source: ${foundSourceSheet}`);

                // 4. Resolve Source Sheet to ID using allSheetsContext
                // We depend on allSheetsContext being populated by loadSheets
                const targetContext = allSheetsContext.find(ctx => ctx.sheetName.toLowerCase() === foundSourceSheet.toLowerCase());

                if (targetContext) {
                    // Navigate to searchForEmail
                    await searchForEmail(token, targetContext.spreadsheetId, targetContext.sheetName, email);
                } else {
                    showStatus(`Looked for "${foundSourceSheet}" but could not find it in loaded sheets.`, 'error');
                }

            } else {
                showStatus('Email not found in the last 1500 entries of SentLog.', 'error');
            }

        } catch (err) {
            console.error(err);
            showStatus(`SentLog Search Error: ${err.message}`, 'error');
        }
    }

    async function searchAllSheets(token, email) {
        showStatus('Searching all sheets...');
        resultSection.classList.add('hidden');

        let found = false;

        for (const ctx of allSheetsContext) {
            showStatus(`Searching in ${ctx.sourceName}: ${ctx.sheetName}...`);
            try {
                // Reuse the logic but return true/false instead of void 
                // We'll refactor searchForEmail to be return-based or split the logic.
                // For minimally invasive change, let's just copy the fetch/check logic here tailored for loop.

                const range = `${ctx.sheetName}!A:Z`;
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${ctx.spreadsheetId}/values/${encodeURIComponent(range)}`;

                const response = await fetch(url, {
                    headers: { 'Authorization': `Bearer ${token}` }
                });

                if (!response.ok) continue; // Skip if error on one sheet

                const data = await response.json();
                if (!data.values || data.values.length === 0) continue;

                const headers = data.values[0];
                const emailIndex = headers.findIndex(h => h && h.trim().toLowerCase() === 'email');
                if (emailIndex === -1) continue;

                const searchEmail = email.toLowerCase().trim();
                let matchRow = -1;
                let matchData = null;

                for (let i = 1; i < data.values.length; i++) {
                    const row = data.values[i];
                    const rowEmail = (row[emailIndex]) ? String(row[emailIndex]).toLowerCase().trim() : '';
                    if (rowEmail === searchEmail) {
                        matchRow = i;
                        matchData = row;
                        break;
                    }
                }

                if (matchRow !== -1) {
                    // FOUND!
                    // Call the display logic (we can reuse searchForEmail or extract display logic)
                    // Simplest is to just call searchForEmail knowing it will succeed and handle UI
                    await searchForEmail(token, ctx.spreadsheetId, ctx.sheetName, email);
                    found = true;
                    break;
                }

            } catch (ignore) {
                console.error(ignore);
            }
        }

        if (!found) {
            showStatus('Email not found in any sheet.', 'error');
        }
    }

    async function searchForEmail(token, spreadsheetId, sheetName, email) {
        showStatus(`Searching...`);
        resultSection.classList.add('hidden');
        currentSpreadsheetId = spreadsheetId;
        currentSheetName = sheetName;

        try {
            const range = `${sheetName}!A:Z`;
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;

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

                // Pre-fill Calendly inputs
                calNameInput.value = (name === '(Name not found)') ? '' : name;
                calEmailInput.value = matchData[emailIndex];

                const currentStatus = (matchData[respondedIndex]) ? String(matchData[respondedIndex]).trim() : '';

                let optionFound = false;
                // Try to find exact match
                for (let opts of responseSelect.options) {
                    if (opts.value === currentStatus) {
                        responseSelect.value = currentStatus;
                        optionFound = true;
                        break;
                    }
                }
                // If not found, try case-insensitive match
                if (!optionFound && currentStatus) {
                    const currentLower = currentStatus.toLowerCase();
                    for (let opts of responseSelect.options) {
                        if (opts.value.toLowerCase() === currentLower) {
                            responseSelect.value = opts.value;
                            optionFound = true;
                            break;
                        }
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

    async function updateResponseStatus(token, spreadsheetId, sheetName, rowIndex, newStatus) {
        if (!currentUpdateRange) return;

        showStatus('Updating...');
        try {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(currentUpdateRange)}?valueInputOption=USER_ENTERED`;

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

    // --- Calendly Helper Functions ---

    async function initCalendly(token) {
        try {
            // 1. Get User URI
            const userRes = await fetch('https://api.calendly.com/users/me', {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            if (!userRes.ok) throw new Error('Failed to fetch Calendly user');
            const userData = await userRes.json();
            calendlyUserUri = userData.resource.uri;

            // 2. Get Event Types
            const eventsRes = await fetch(`https://api.calendly.com/event_types?user=${calendlyUserUri}`, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            if (!eventsRes.ok) throw new Error('Failed to fetch Event Types');
            const eventsData = await eventsRes.json();

            eventTypeSelect.innerHTML = '<option value="" disabled selected>Select event type...</option>';
            eventsData.collection.forEach(et => {
                if (et.active) {
                    const opt = document.createElement('option');
                    opt.value = et.uri;
                    opt.textContent = et.name;
                    eventTypeSelect.appendChild(opt);
                }
            });

            calendlySection.classList.remove('hidden');

        } catch (err) {
            console.error(err);
            showStatus('Calendly Init Failed: ' + err.message, 'error');
            tokenInputSection.classList.remove('hidden'); // Show input to correct it
        }
    }

    async function createCalendlyLink(token, eventTypeUri, name, email) {
        showStatus('Generating link...');
        calendlyResult.classList.add('hidden');
        try {
            const response = await fetch('https://api.calendly.com/scheduling_links', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    max_event_count: 1,
                    owner: eventTypeUri,
                    owner_type: "EventType"
                })
            });

            if (!response.ok) {
                const errText = await response.text();
                throw new Error(errText);
            }

            const data = await response.json();
            let link = data.resource.booking_url;

            // Pre-fill parameters
            const params = new URLSearchParams();
            if (name) params.append('name', name);
            if (email) params.append('email', email);

            // Check if link already has query params? Usually create_link returns a clean URL.
            // But just in case, logic could be cleaner. 
            // Standard Calendly links don't usually have params unless custom.

            if (Array.from(params).length > 0) {
                link += (link.includes('?') ? '&' : '?') + params.toString();
            }

            generatedLinkDisplay.value = link;
            calendlyResult.classList.remove('hidden');
            showStatus('Link generated!', 'success');

        } catch (err) {
            showStatus('Link Gen Failed: ' + err.message, 'error');
        }
    }
});
