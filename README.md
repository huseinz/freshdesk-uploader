<!--
    VAB Freshdesk Uploader
    This HTML file allows users to upload contacts and tickets to Freshdesk using their API.
    Users can select an XLSX file, choose a sheet, and upload the data.
    The script handles both creating and updating contacts based on the provided data.
    The same functionality is available for tickets as well.

    This is implemented as a POC and may require further enhancements for production use.
    -- Zubir Husein, 2025 <zubirhusein@gmail.com>
-->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Freshdesk Integration</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js"></script>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        input, button { margin: 5px 0; padding: 10px; }
        #output { margin-top: 20px; }
        .message { margin: 10px 0; padding: 10px; border: 1px solid #ccc; }
        .success { background-color: #e0ffe0; border-color: #00cc00; }
        .error { background-color: #ffe0e0; border-color: #cc0000; }
        /* Apply a 1px solid black border to the table, header cells, and data cells */
        table, th, td {
          border: 1px solid black;
        }
        /* Collapse double borders into single borders for a cleaner look */
        table {
          border-collapse: collapse;
        }
        th {
          background-color: #f2f2f2; /* Light gray background for headers */
          border-bottom: 2px solid #ccc; /* Thicker bottom border for headers */
        }
        td {
          border-right: 1px dotted gray; /* Dotted right border for data cells */
          white-space: nowrap;
        }
    </style>
</head>
<body>
    <h1>VAB Freshdesk Uploader</h1>
    <label for="apiKey">API Key:</label><br>
    <input type="password" id="apiKey" placeholder="Profile settings > View API Key" size="40"><br>
    <div id="ratelimit"></div>
    <button id="listContactFields">List Contact Fields</button>
    <button id="listTicketFields">List Ticket Fields</button><br><br>
    <input type="file" id="contactFileInput" accept=".xlsx"><br>
    <button id="uploadContacts">Upload Contacts</button>
    <select id="contactSheetDropdown"></select><br><br>
    <input type="file" id="ticketFileInput" accept=".xlsx"><br>
    <button id="uploadTickets">Upload Tickets</button>
    <select id="ticketSheetDropdown"></select>
    <div id="progress"></div>
    <div id="status"></div>
    <div id="tableContainer"></div>
    <script>
        // Configuration
        const domain = 'nightingale-support.freshdesk.com'
        const progress = document.getElementById('progress');
        const rateLimit = document.getElementById('ratelimit');
        const status = document.getElementById('status');
        const table = document.getElementById('tableContainer');
        const contactSheetDropdown = document.getElementById('contactSheetDropdown');

        class Freshdesk {
            static previousHeaders = null; // Store previous headers for rate limit comparison
            constructor(apiKey) {
                this.apiKey = apiKey;
                this.headers = {
                    'Authorization': 'Basic ' + btoa(apiKey + ':X'),
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': '*'
                };
            }
            async make_request(endpoint, method='GET', body=null) {
                await displayRateLimit(Freshdesk.previousHeaders);
                let fetchOptions = {
                    method: method,
                    headers: this.headers
                };
                if (body) {
                    fetchOptions.body = JSON.stringify(body);
                }
                let response = await fetch(`https://${domain}${endpoint}`, fetchOptions);
                // handle 401
                if (response.status === 401){
                    throw new Error('Invalid API key!');
                }
                Freshdesk.previousHeaders = response.headers; // Store current headers for next request
                return response;
            }
            async listContactFields() {
                let response = await this.make_request(`/api/v2/contact_fields`, 'GET');
                if (!response.ok) {
                    throw new Error(`Error fetching contact fields: ${response.statusText}`);
                }
                return await response.json();
            }
            async listTicketFields() {
                let response = await this.make_request(`/api/v2/ticket_fields`, 'GET');
                if (!response.ok) {
                    throw new Error(`Error fetching ticket fields: ${response.statusText}`);
                }
                return await response.json();
            }

            // Upsert contact to Freshdesk
            async upsertContact(contact) {
                let formatted_contact = {'custom_fields': {}};
                // Contact field mappings
                formatted_contact.unique_external_id = contact['Medicaid No'];
                formatted_contact.mobile = contact['Mobile Phone Number'];
                formatted_contact.description = '<pre>' + contact['COMMENTS'] + '</pre>';
                formatted_contact.custom_fields.issue_id = contact['Issue Id'];
                formatted_contact.custom_fields.subscribermember_id = contact['Subscriber Id'];
                formatted_contact.custom_fields.medicare_no = contact['Medicare No'];
                formatted_contact.custom_fields.medicaid_no = contact['Medicaid No'];
                formatted_contact.custom_fields.date_of_birth = contact['Date of Birth'];
                formatted_contact.custom_fields.address_line_1 = contact['Address Line 1'];
                formatted_contact.custom_fields.address_line_2 = contact['Address Line 2'];
                formatted_contact.custom_fields.city = contact['City'];
                formatted_contact.custom_fields.state = contact['State'];
                formatted_contact.custom_fields.zip_code = contact['Zip Code'];
                formatted_contact.custom_fields.county = contact['County'];
                formatted_contact.custom_fields.effective_date = contact['Effective Date'];
                formatted_contact.custom_fields.plan_code = contact['Plan Code'];
                formatted_contact.custom_fields.home_phone = contact['Home Phone Number'];

                // Combine first_name and last_name into name
                const firstName = contact['First Name'] || '';
                const lastName = contact['Last Name'] || '';
                formatted_contact.name = (firstName + ' ' + lastName).trim();
                // combine street, city, state, zip_code into address
                const addressParts = [];
                if (formatted_contact.custom_fields.address_line_1) addressParts.push(formatted_contact.custom_fields.address_line_1);
                if (formatted_contact.custom_fields.address_line_2) addressParts.push(formatted_contact.custom_fields.address_line_2);
                if (formatted_contact.custom_fields.city) addressParts.push(formatted_contact.custom_fields.city);
                if (formatted_contact.custom_fields.state) addressParts.push(formatted_contact.custom_fields.state);
                if (formatted_contact.custom_fields.zip_code) addressParts.push(formatted_contact.custom_fields.zip_code);
                if (addressParts.length > 0) {
                    formatted_contact.custom_fields.full_mailing_address = addressParts.join(' ');
                }

                // Attempt to create the contact
                console.log('Creating contact:', formatted_contact);
                let response = await this.make_request(`/api/v2/contacts`, 'POST', formatted_contact);
                if (response.ok) {
                    return 'Success (Created New Contact)';
                }
                let returnData = await response.json();
                console.log('Error data:', returnData);

                // Check if error is due to existing contact
                let user_id = null;
                for (const error of returnData.errors) {
                    if (error.additional_info && error.additional_info.user_id) {
                        user_id = error.additional_info.user_id;
                        break;
                    }
                }
                // If no existing user_id found, rethrow the error
                if (!user_id) {
                    throw new Error(`Error creating contact: ${JSON.stringify(returnData.errors)}`);
                }

                // Contact with this ID already exists, update it
                console.log('Existing User ID?:', user_id);
                response = await this.make_request(`/api/v2/contacts/${user_id}`, 'PUT', formatted_contact);
                if (!response.ok) {
                    let returnData = await response.json();
                    throw new Error(`Error updating contact: ${JSON.stringify(returnData.errors)}`);
                }
                return 'Success (Updated Existing Contact)';
            }

            async createTicket(ticket) {
                let formatted_ticket = {'custom_fields': {}};
                // Ticket field mappings
                formatted_ticket.status = 2; // Open
                formatted_ticket.priority = 2; // Low
                formatted_ticket.type = 'Escalation';
                formatted_ticket.unique_external_id = ticket['Medicaid Id'];
                formatted_ticket.description = '<pre>' + ticket['Ticket notes'] + '</pre>';
                formatted_ticket.subject = ticket['Medicaid Id'] + ': ' + ticket['Member name'] + ' - ' + ticket['Reason For Escalation'];
                formatted_ticket.custom_fields.cf_careconnect_ticket_no = ticket['Careconnect Ticket No'];
                formatted_ticket.custom_fields.cf_initial_request_date = ticket['Initial Request Date'];
                formatted_ticket.custom_fields.cf_wellcare_id_number = ticket['WellCare ID Number'];
                formatted_ticket.custom_fields.cf_address = ticket['Address'];
                formatted_ticket.custom_fields.cf_medicaid_id = ticket['Medicaid Id'];
                formatted_ticket.custom_fields.cf_phone_number = ticket['Phone number'];
                formatted_ticket.custom_fields.cf_caller_phone_number = ticket['Caller Phone Number'];
                formatted_ticket.custom_fields.cf_utility_company_name = ticket['Utility Company Name'];
                formatted_ticket.custom_fields.cf_utility_account_number = ticket['Utility Account Number'];
                formatted_ticket.custom_fields.cf_leasing_agent_name = ticket['Leasing Agent Name'];
                formatted_ticket.custom_fields.cf_leasing_agent_phone = ticket['Leasing Agent Phone'];

                // Make request to create ticket
                console.log('Creating ticket:', formatted_ticket);
                let response = await this.make_request(`/api/v2/tickets`, 'POST', formatted_ticket);
                if (response.ok) {
                    let returnData = await response.json();
                    console.log('Created ticket:', returnData);
                    return 'Success (Created New Ticket)';
                }
                let returnData = await response.json();
                throw new Error(`Error creating ticket: ${JSON.stringify(returnData.errors)}`);
            }
        }
        function sleep(ms) {
            return new Promise(resolve => setTimeout(resolve, ms));
        }
        // Clear existing content
        function clearAll() {
            progress.textContent = '';
            status.innerHTML = '';
            table.innerHTML = '';
        }
        // Display rate limit
        async function displayRateLimit(previousHeaders) {
            if (!previousHeaders) {
                return;
            }
            let limit = previousHeaders.get('X-RateLimit-Total');
            let remaining = previousHeaders.get('X-RateLimit-Remaining');
            rateLimit.className = 'message success';
            // calculate how much to sleep on average in between every request to avoid hitting the limit if limit resets every minute
            const resetInterval = 60000; // 1 minute in milliseconds
            const averageSleepTime = resetInterval / remaining;
            rateLimit.textContent = `Rate Limit: ${limit}, Remaining: ${remaining}, waiting ${Math.round(averageSleepTime)} ms between requests`;
            await sleep(averageSleepTime);
        }
        // Display messages
        function displayMessage(message, type) {
            const div = document.createElement('div');
            div.className = `message ${type}`;
            div.textContent = message;
            status.appendChild(div);
        }
        // Display progress
        function displayProgress(message) {
            progress.className = 'message success';
            progress.textContent = message;
        }
        // Create dynamic table
        function blacklistColumns(row) {
            const keysToExclude = ['Notes', 'Ticket notes', 'COMMENTS'];
            return Object.fromEntries(Object.entries(row).filter(([key]) => !keysToExclude.includes(key)));
        }
        function initDynamicTable(data, container) {
            if (!container) {
                console.error("Container element not found.");
                return;
            }

            const table = document.createElement("table");
            const thead = document.createElement("thead");
            const tbody = document.createElement("tbody");

            // Create table headers
            const rowData = blacklistColumns(data);
            // Create header row dynamically based on keys of the first object
            const headerRow = document.createElement("tr");
            Object.keys(rowData).forEach(key => {
                const th = document.createElement("th");
                th.textContent = key.charAt(0).toUpperCase() + key.slice(1); // Capitalize first letter
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);

            table.appendChild(thead);
            table.appendChild(tbody);
            container.appendChild(table);
            return tbody;
        }
        function updateDynamicTable(data, container) {
            if (!container) {
                console.error("Container element not found.");
                return;
            }
            const rowData = blacklistColumns(data);
            // Create table rows and cells
            const row = document.createElement("tr");
            row.classList.add('error');
            Object.values(rowData).forEach(value => {
                const td = document.createElement("td");
                td.textContent = value;
                if (typeof value === 'string' && value.startsWith('Success')) {
                    row.classList.add('success');
                    row.classList.remove('error');
                }
                row.appendChild(td);
            });
            container.appendChild(row);
        }
        // load and display sheets in dropdown
        async function displaySheetsInDropdown(fileInputEvent, dropdownId) {
            const fileInput = fileInputEvent.target;
            const dropdownElement = document.getElementById(dropdownId);
            if (fileInput.files.length === 0) {
                return;
            }
            const file = fileInput.files[0];
            const sheetNames = await fetchSheets(file);
            dropdownElement.innerHTML = '';
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = 'Select a sheet';
            dropdownElement.appendChild(defaultOption);
            sheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                dropdownElement.appendChild(option);
            });
        }
        // Fetch sheet names from XLSX file
        async function fetchSheets(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', bookSheets: true });
                    resolve(workbook.SheetNames);
                };
                reader.onerror = (error) => reject(error);
                reader.readAsArrayBuffer(file);
            });
        }

        /*
        Contact Upload Logic
        */
        // Upload contacts from XLSX file
        async function uploadContacts() {
            clearAll();
            const apiKey = document.getElementById('apiKey').value;
            const fileInput = document.getElementById('contactFileInput');
            if (!apiKey) {
                displayMessage('Please enter your API key.', 'error');
                return;
            }
            if (fileInput.files.length === 0) {
                displayMessage('Please select an XLSX file.', 'error');
                return;
            }
            const contactSheetName = contactSheetDropdown.value;
            if (!contactSheetName) {
                displayMessage('Please select a sheet from the dropdown.', 'error');
                return;
            }
            const freshdesk = new Freshdesk(apiKey);

            const file = fileInput.files[0];
            const reader = new FileReader();
            reader.onload = async (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type:'array',cellText:false,cellDates:true});
                const sheet = workbook.Sheets[contactSheetName];
                if (!sheet) {
                    displayMessage(`Sheet "${contactSheetName}" not found in the file.`, 'error');
                    return;
                }
                const contacts = XLSX.utils.sheet_to_json(sheet, {raw:false,dateNF:'yyyy-mm-dd', defval: null});
                if (contacts.length === 0) {
                    displayMessage('The selected file is empty or not properly formatted.', 'error');
                    return;
                }
                displayProgress(`Loaded ${contacts.length} contact(s) from the file. Starting upload...`);
                let successCount = 0;
                let errorCount = 0;
                let processedCount = 0;
                contacts[0].status = '';
                const tbody = initDynamicTable(contacts[0], table);
                for (const contact of contacts) {
                    displayProgress(`Uploading contact ${processedCount + 1} of ${contacts.length}...`);
                    processedCount++;
                    try {
                        const response = await freshdesk.upsertContact(contact);
                        successCount++;
                        contact.status = response;
                    } catch (error) {
                        errorCount++;
                        console.log('Failed response:', error.message);
                        contact.status = error.message;
                    }
                    updateDynamicTable(contact, tbody);
                }
                if (successCount > 0) {
                    displayMessage(`Successfully uploaded ${successCount} out of ${contacts.length} contact(s).`, 'success');
                }
                if (errorCount > 0) {
                    displayMessage(`Failed to upload ${errorCount} contact(s). See below table for details.`, 'error');
                }
                displayProgress('Upload process completed.');
            };
            reader.readAsArrayBuffer(file);
        }

        async function displayContactFields() {
            clearAll();
            const apiKey = document.getElementById('apiKey').value;
            if (!apiKey) {
                displayMessage('Please enter your API key.', 'error');
                return;
            }
            const freshdesk = new Freshdesk(apiKey);
            try {
                const fields = await freshdesk.listContactFields();
                if (fields.length === 0) {
                    displayMessage('No contact fields found.', 'error');
                    return;
                }
                const tbody = initDynamicTable(fields[0], table);
                fields.forEach(field => updateDynamicTable(field, tbody));
                displayMessage(`Fetched ${fields.length} contact field(s).`, 'success');
            } catch (error) {
                displayMessage(`Error fetching contact fields: ${error.message}`, 'error');
            }
        }
        async function displayTicketFields() {
            clearAll();
            const apiKey = document.getElementById('apiKey').value;
            if (!apiKey) {
                displayMessage('Please enter your API key.', 'error');
                return;
            }
            const freshdesk = new Freshdesk(apiKey);
            try {
                const fields = await freshdesk.listTicketFields();
                if (fields.length === 0) {
                    displayMessage('No ticket fields found.', 'error');
                    return;
                }
                const tbody = initDynamicTable(fields[0], table);
                fields.forEach(field => updateDynamicTable(field, tbody));
                displayMessage(`Fetched ${fields.length} ticket field(s).`, 'success');
            } catch (error) {
                displayMessage(`Error fetching ticket fields: ${error.message}`, 'error');
            }
        }

        async function uploadTickets() {
            clearAll();
            const apiKey = document.getElementById('apiKey').value;
            const fileInput = document.getElementById('ticketFileInput');
            const ticketSheetDropdown = document.getElementById('ticketSheetDropdown');
            if (!apiKey) {
                displayMessage('Please enter your API key.', 'error');
                return;
            }
            if (fileInput.files.length === 0) {
                displayMessage('Please select an XLSX file.', 'error');
                return;
            }
            const ticketSheetName = ticketSheetDropdown.value;
            if (!ticketSheetName) {
                displayMessage('Please select a sheet from the dropdown.', 'error');
                return;
            }
            const freshdesk = new Freshdesk(apiKey);

            const file = fileInput.files[0];
            const reader = new FileReader();
            reader.onload = async (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type:'array',cellText:false,cellDates:true});
                const sheet = workbook.Sheets[ticketSheetName];
                if (!sheet) {
                    displayMessage(`Sheet "${ticketSheetName}" not found in the file.`, 'error');
                    return;
                }
                const tickets = XLSX.utils.sheet_to_json(sheet, {raw:false,dateNF:'yyyy-mm-dd', defval: null});
                if (tickets.length === 0) {
                    displayMessage('The selected file is empty or not properly formatted.', 'error');
                    return;
                }
                displayProgress(`Loaded ${tickets.length} ticket(s) from the file. Starting upload...`);
                let successCount = 0;
                let errorCount = 0;
                let processedCount = 0;
                tickets[0].status = '';
                const tbody = initDynamicTable(tickets[0], table);
                for (const ticket of tickets) {
                    displayProgress(`Uploading ticket ${processedCount + 1} of ${tickets.length}...`);
                    processedCount++;
                    try {
                        const response = await freshdesk.createTicket(ticket);
                        successCount++;
                        ticket.status = response;
                    } catch (error) {
                        errorCount++;
                        console.log('Failed response:', error.message);
                        ticket.status = error.message;
                    }
                    updateDynamicTable(ticket, tbody);
                }
                if (successCount > 0) {
                    displayMessage(`Successfully uploaded ${successCount} out of ${tickets.length} ticket(s).`, 'success');
                }
                if (errorCount > 0) {
                    displayMessage(`Failed to upload ${errorCount} ticket(s). See below table for details.`, 'error');
                }
                displayProgress('Upload process completed.');
            };
            reader.readAsArrayBuffer(file);
        }

        document.getElementById('contactFileInput').addEventListener('change', (e) => displaySheetsInDropdown(e, 'contactSheetDropdown'));
        document.getElementById('uploadContacts').addEventListener('click', uploadContacts);
        document.getElementById('ticketFileInput').addEventListener('change', (e) => displaySheetsInDropdown(e, 'ticketSheetDropdown'));
        document.getElementById('uploadTickets').addEventListener('click', uploadTickets);
        document.getElementById('listContactFields').addEventListener('click', displayContactFields);
        document.getElementById('listTicketFields').addEventListener('click', displayTicketFields);
    </script>
</body>
</html>
