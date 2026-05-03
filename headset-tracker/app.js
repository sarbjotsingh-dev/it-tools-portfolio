// Headset Management System
// SharePoint Integration Script

(function() {
    // SharePoint Configuration
    const LIST_GUID = "5074921e-be56-4d84-9750-bfbed8375e8c";
    const SITE_URL = "https://your-tenant.sharepoint.com";
    
    // Hardcoded Admin Credentials
    const ADMIN_USERNAME = "admin";
    const ADMIN_PASSWORD = "Light44@@";
    
    // Global variables
    let masterData = [];
    let csvRows = [];
    let requestDigest = "";
    let isAdminLoggedIn = false;
    let currentEditingAssetId = null;
    
    // Choice options
    let choiceOptions = {
        OfficeLocation: [],
        Department: [],
        Status: ["Available", "Assigned", "Under Repair", "Retired"],
        HeadsetType: []
    };
    
    // Normalization maps
    let normalizedChoiceMaps = {};
    
    // DOM Elements
    let elements = {};
    
    // Initialize application
    async function init() {
        cacheElements();
        attachEventListeners();
        ensureDatalists();
        await loadChoiceOptions();
        await loadTenantUsers();
        await loadData();
        await refreshDigest();
        populateDropdowns();
        updateSearchSuggestions();
        updateUserSuggestions();
    }
    
    function ensureDatalists() {
        if (!elements.assetIdSuggestions) {
            const datalist = document.createElement('datalist');
            datalist.id = 'assetIdSuggestions';
            document.body.appendChild(datalist);
            elements.assetIdSuggestions = datalist;
        }
        if (!elements.serialSuggestions) {
            const datalist = document.createElement('datalist');
            datalist.id = 'serialSuggestions';
            document.body.appendChild(datalist);
            elements.serialSuggestions = datalist;
        }
        if (!elements.assignToNameSuggestions) {
            const datalist = document.createElement('datalist');
            datalist.id = 'assignToNameSuggestions';
            document.body.appendChild(datalist);
            elements.assignToNameSuggestions = datalist;
        }
        if (elements.searchAssetId) {
            elements.searchAssetId.setAttribute('list', 'assetIdSuggestions');
        }
        if (elements.searchSerial) {
            elements.searchSerial.setAttribute('list', 'serialSuggestions');
        }
    }
    
    function updateSearchSuggestions() {
        if (!elements.assetIdSuggestions || !elements.serialSuggestions) return;

        const assetIds = Array.from(new Set(masterData.map(item => (item.AssetID || '').trim()).filter(Boolean))).slice(0, 1000);
        const serials = Array.from(new Set(masterData.map(item => (item.SerialNumber || '').trim()).filter(Boolean))).slice(0, 1000);

        elements.assetIdSuggestions.innerHTML = assetIds.map(val => `<option value="${escapeHtml(val)}"></option>`).join('');
        elements.serialSuggestions.innerHTML = serials.map(val => `<option value="${escapeHtml(val)}"></option>`).join('');
    }

    async function loadTenantUsers() {
        try {
            const response = await fetch(`${SITE_URL}/_api/web/siteusers?$select=Title,Email&$top=1000`, {
                headers: { "Accept": "application/json;odata=nometadata" }
            });
            const data = await response.json();
            const userItems = data.value || [];
            window.tenantUsers = userItems.map(u => (u.Title || u.Email || '').trim()).filter(Boolean);
            updateUserSuggestions();
        } catch (e) {
            console.error("Error loading tenant users:", e);
            window.tenantUsers = [];
        }
    }

    function updateUserSuggestions() {
        if (!elements.assignToNameSuggestions) return;
        const users = Array.from(new Set((window.tenantUsers || []).map(u => u.trim()).filter(Boolean))).slice(0, 1000);
        elements.assignToNameSuggestions.innerHTML = users.map(u => `<option value="${escapeHtml(u)}"></option>`).join('');
    }

    // Cache DOM elements for better performance
    function cacheElements() {
        elements = {
            adminLoginBtn: document.getElementById('adminLoginBtn'),
            adminView: document.getElementById('adminView'),
            userView: document.getElementById('userView'),
            backToUserBtn: document.getElementById('backToUserBtn'),
            loginModal: document.getElementById('loginModal'),
            loginSubmitBtn: document.getElementById('loginSubmitBtn'),
            loginUsername: document.getElementById('loginUsername'),
            loginPassword: document.getElementById('loginPassword'),
            searchBtn: document.getElementById('searchBtn'),
            searchAssetId: document.getElementById('searchAssetId'),
            searchSerial: document.getElementById('searchSerial'),
            resultSection: document.getElementById('resultSection'),
            resultTitle: document.getElementById('resultTitle'),
            resultContent: document.getElementById('resultContent'),
            addFormSection: document.getElementById('addFormSection'),
            addHeadsetBtn: document.getElementById('addHeadsetBtn'),
            cancelAddBtn: document.getElementById('cancelAddBtn'),
            newAssetId: document.getElementById('newAssetId'),
            newSerial: document.getElementById('newSerial'),
            newOfficeLocation: document.getElementById('newOfficeLocation'),
            newDepartment: document.getElementById('newDepartment'),
            newHeadsetType: document.getElementById('newHeadsetType'),
            newPersonAssigned: document.getElementById('newPersonAssigned'),
            tabView: document.getElementById('tabView'),
            tabBulk: document.getElementById('tabBulk'),
            viewSection: document.getElementById('viewSection'),
            bulkSection: document.getElementById('bulkSection'),
            fAssetId: document.getElementById('fAssetId'),
            fSerial: document.getElementById('fSerial'),
            fStatus: document.getElementById('fStatus'),
            fType: document.getElementById('fType'),
            mainTbody: document.getElementById('mainTbody'),
            downloadTemplateBtn: document.getElementById('downloadTemplateBtn'),
            csvFile: document.getElementById('csvFile'),
            startImportBtn: document.getElementById('startImportBtn'),
            csvSummary: document.getElementById('csvSummary'),
            csvPreviewBody: document.getElementById('csvPreviewBody'),
            csvLog: document.getElementById('csvLog'),
            statusBox: document.getElementById('statusBox'),
            totalCount: document.getElementById('totalCount'),
            assignedCount: document.getElementById('assignedCount'),
            availableCount: document.getElementById('availableCount'),
            repairCount: document.getElementById('repairCount'),
            verifyModal: document.getElementById('verifyModal'),
            verifySaveBtn: document.getElementById('verifySaveBtn'),
            verifyCancelBtn: document.getElementById('verifyCancelBtn'),
            verifyAssetId: document.getElementById('verifyAssetId'),
            verifySerial: document.getElementById('verifySerial'),
            verifyOfficeLocation: document.getElementById('verifyOfficeLocation'),
            verifyDepartment: document.getElementById('verifyDepartment'),
            verifyHeadsetType: document.getElementById('verifyHeadsetType'),
            verifyStatus: document.getElementById('verifyStatus'),
            verifyPersonAssigned: document.getElementById('verifyPersonAssigned'),
            verifyNotes: document.getElementById('verifyNotes')
        };
    }
    
    // Attach all event listeners
    function attachEventListeners() {
        if (elements.adminLoginBtn) {
            elements.adminLoginBtn.addEventListener('click', showAdminLogin);
        }
        
        if (elements.loginSubmitBtn) {
            elements.loginSubmitBtn.addEventListener('click', checkAdminLogin);
        }
        
        if (elements.backToUserBtn) {
            elements.backToUserBtn.addEventListener('click', logoutAdmin);
        }
        
        if (elements.searchBtn) {
            elements.searchBtn.addEventListener('click', searchHeadset);
        }
        
        if (elements.addHeadsetBtn) {
            elements.addHeadsetBtn.addEventListener('click', addNewHeadset);
        }
        
        if (elements.cancelAddBtn) {
            elements.cancelAddBtn.addEventListener('click', clearAddForm);
        }
        
        if (elements.tabView) {
            elements.tabView.addEventListener('click', () => switchTab('view'));
        }
        
        if (elements.tabBulk) {
            elements.tabBulk.addEventListener('click', () => switchTab('bulk'));
        }
        
        if (elements.downloadTemplateBtn) {
            elements.downloadTemplateBtn.addEventListener('click', downloadTemplate);
        }
        
        if (elements.startImportBtn) {
            elements.startImportBtn.addEventListener('click', importCsvRows);
        }
        
        if (elements.csvFile) {
            elements.csvFile.addEventListener('change', handleCSVFile);
        }
        
        if (elements.fAssetId) {
            elements.fAssetId.addEventListener('input', renderAdminTable);
        }
        
        if (elements.fSerial) {
            elements.fSerial.addEventListener('input', renderAdminTable);
        }
        
        if (elements.fStatus) {
            elements.fStatus.addEventListener('change', renderAdminTable);
        }
        
        if (elements.fType) {
            elements.fType.addEventListener('change', renderAdminTable);
        }
        
        if (elements.verifySaveBtn) {
            elements.verifySaveBtn.addEventListener('click', saveVerifyChanges);
        }

        if (elements.verifyCancelBtn) {
            elements.verifyCancelBtn.addEventListener('click', closeVerifyModal);
        }
        
        // Enter key search
        const searchInputs = [elements.searchAssetId, elements.searchSerial];
        searchInputs.forEach(input => {
            if (input) {
                input.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') {
                        searchHeadset();
                    }
                });
            }
        });
    }
    
    // Load choice options from SharePoint
    async function loadChoiceOptions() {
        for (const field of Object.keys(choiceOptions)) {
            if (field === 'Status') continue;
            try {
                const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/fields?$filter=EntityPropertyName eq '${field}'`, {
                    headers: { "Accept": "application/json;odata=nometadata" }
                });
                const data = await response.json();
                if (data.value && data.value[0] && data.value[0].Choices) {
                    choiceOptions[field] = data.value[0].Choices;
                    createNormalizationMap(field, data.value[0].Choices);
                }
            } catch (e) {
                console.error(`Error loading ${field}:`, e);
            }
        }
    }
    
    // Create normalization map for fuzzy matching
    function createNormalizationMap(fieldName, choices) {
        const map = {};
        choices.forEach(choice => {
            const normalized = choice.toLowerCase().replace(/[^a-z0-9]/g, '');
            map[normalized] = choice;
            
            if (fieldName === 'HeadsetType') {
                if (choice.toLowerCase().includes('mono')) map['mono'] = choice;
                if (choice.toLowerCase().includes('stereo')) map['stereo'] = choice;
                if (choice.toLowerCase().includes('usb')) map['usb'] = choice;
            }
        });
        normalizedChoiceMaps[fieldName] = map;
    }
    
    // Find matching choice
    function findMatchingChoice(fieldName, userInput) {
        if (!userInput || userInput.trim() === '') return '';
        
        const input = userInput.trim();
        const choices = choiceOptions[fieldName] || [];
        
        const exactMatch = choices.find(c => c.toLowerCase() === input.toLowerCase());
        if (exactMatch) return exactMatch;
        
        const normalizedInput = input.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (normalizedChoiceMaps[fieldName] && normalizedChoiceMaps[fieldName][normalizedInput]) {
            return normalizedChoiceMaps[fieldName][normalizedInput];
        }
        
        const partialMatch = choices.find(c => c.toLowerCase().includes(input.toLowerCase()));
        if (partialMatch) return partialMatch;
        
        return input;
    }
    
    // Populate dropdowns
    function populateDropdowns() {
        if (elements.newOfficeLocation) {
            elements.newOfficeLocation.innerHTML = '<option value="">Select Office Location</option>';
            choiceOptions.OfficeLocation.forEach(opt => {
                elements.newOfficeLocation.innerHTML += `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`;
            });
        }
        
        if (elements.newDepartment) {
            elements.newDepartment.innerHTML = '<option value="">Select Department</option>';
            choiceOptions.Department.forEach(opt => {
                elements.newDepartment.innerHTML += `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`;
            });
        }
        
        if (elements.newHeadsetType) {
            elements.newHeadsetType.innerHTML = '<option value="">Select Headset Type</option>';
            choiceOptions.HeadsetType.forEach(opt => {
                elements.newHeadsetType.innerHTML += `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`;
            });
            const defaultOption = choiceOptions.HeadsetType.find(opt => opt.toLowerCase().includes('logitech stereo usb')) || choiceOptions.HeadsetType[0];
            if (defaultOption) elements.newHeadsetType.value = defaultOption;
        }
        if (elements.verifyHeadsetType) {
            elements.verifyHeadsetType.innerHTML = '<option value="">Select Headset Type</option>';
            choiceOptions.HeadsetType.forEach(opt => {
                elements.verifyHeadsetType.innerHTML += `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`;
            });
        }
    }
    
    // Refresh SharePoint digest token
    async function refreshDigest() {
        try {
            const response = await fetch(`${SITE_URL}/_api/contextinfo`, {
                method: "POST",
                headers: { "Accept": "application/json;odata=nometadata" }
            });
            const data = await response.json();
            requestDigest = data.FormDigestValue;
            setTimeout(refreshDigest, 1500000);
        } catch (e) {
            setTimeout(refreshDigest, 300000);
        }
    }
    
    // Load data from SharePoint
    async function loadData() {
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items?$select=Id,AssetID,OfficeLocation,Department,SerialNumber,Status,PersonAssigned,HeadsetType,Notes&$top=5000`, {
                headers: { "Accept": "application/json;odata=nometadata" }
            });
            const data = await response.json();
            masterData = data.value || [];
            updateSearchSuggestions();
            updateUserSuggestions();
            
            if (isAdminLoggedIn) {
                updateStatistics();
                renderAdminTable();
            }
        } catch (e) {
            console.error("Load error:", e);
            showStatus("Error loading data", "error");
        }
    }
    
    // Search headset
    async function searchHeadset() {
        const assetId = elements.searchAssetId.value.trim();
        const serial = elements.searchSerial.value.trim();
        
        if (!assetId && !serial) {
            showStatus("Please enter either Asset ID or Serial Number", "warning");
            return;
        }
        
        let foundAsset = null;
        
        if (assetId) {
            foundAsset = masterData.find(a => a.AssetID === assetId);
        }
        
        if (!foundAsset && serial) {
            foundAsset = masterData.find(a => a.SerialNumber === serial);
        }
        
        if (foundAsset) {
            elements.resultTitle.textContent = "✅ Headset Found";
            elements.resultContent.innerHTML = `
                <div class="asset-details">
                    <div class="detail-row">
                        <div class="detail-label">Asset ID:</div>
                        <div>${escapeHtml(foundAsset.AssetID)}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Serial Number:</div>
                        <div>${escapeHtml(foundAsset.SerialNumber || 'N/A')}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Office Location:</div>
                        <div>${escapeHtml(foundAsset.OfficeLocation || 'N/A')}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Department:</div>
                        <div>${escapeHtml(foundAsset.Department || 'N/A')}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Headset Type:</div>
                        <div>${escapeHtml(foundAsset.HeadsetType || 'N/A')}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Current Status:</div>
                        <div class="status-pill">${escapeHtml(foundAsset.Status || 'Unknown')}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Currently Assigned To:</div>
                        <div>
                            ${escapeHtml(foundAsset.PersonAssigned || 'Not Assigned')}
                            <span style="margin-left: 12px;">
                                <input type="text" id="assignToName" list="assignToNameSuggestions" placeholder="Assign to self" style="padding: 6px; width: 180px; margin-right: 6px;"> 
                                <button id="assignHeadsetBtn" class="btn-primary" data-id="${foundAsset.Id}" style="padding: 7px 10px; font-size: 12px;">Assign</button>
                            </span>
                        </div>
                    </div>
                    <div class="assignment-history">
                        <strong>📝 Assignment History:</strong><br>
                        ${escapeHtml(foundAsset.Notes || 'No history available')}
                    </div>
                </div>
                <div style="margin-top: 15px; text-align: right;">
                    <button id="verifyDetailsBtn" class="btn-secondary">Verify / Edit Details</button>
                </div>
            `;
            
            const assignBtn = document.getElementById('assignHeadsetBtn');
            if (assignBtn) {
                assignBtn.addEventListener('click', () => assignHeadset(foundAsset.Id));
            }

            const verifyBtn = document.getElementById('verifyDetailsBtn');
            if (verifyBtn) {
                verifyBtn.addEventListener('click', () => openVerifyModal(foundAsset));
            }
            
            elements.resultSection.classList.add('active');
            elements.addFormSection.classList.add('hidden');
        } else {
            elements.resultTitle.textContent = "❌ Headset Not Found";
            elements.resultContent.innerHTML = `
                <div style="background: #fff3cd; padding: 15px; border-radius: 8px;">
                    <p>No headset found with Asset ID "${escapeHtml(assetId)}" ${serial ? `or Serial Number "${escapeHtml(serial)}"` : ''}</p>
                    <p>Please fill out the form below to add this headset to inventory.</p>
                </div>
            `;
            elements.resultSection.classList.add('active');
            elements.addFormSection.classList.remove('hidden');
            
            if (assetId) elements.newAssetId.value = assetId;
            if (serial) elements.newSerial.value = serial;
            // Default headset type for new records
            if (elements.newHeadsetType && choiceOptions.HeadsetType.length > 0) {
                const defaultType = choiceOptions.HeadsetType.find(opt => opt.toLowerCase().includes('logitech stereo usb')) || choiceOptions.HeadsetType[0];
                if (defaultType) {
                    elements.newHeadsetType.value = defaultType;
                }
            }
        }
    }
    
    function openVerifyModal(asset) {
        currentEditingAssetId = asset.Id;
        if (!elements.verifyModal) return;

        elements.verifyAssetId.value = asset.AssetID || '';
        elements.verifySerial.value = asset.SerialNumber || '';
        elements.verifyOfficeLocation.value = asset.OfficeLocation || '';
        elements.verifyDepartment.value = asset.Department || '';
        elements.verifyHeadsetType.value = asset.HeadsetType || '';
        elements.verifyStatus.value = asset.Status || 'Available';
        elements.verifyPersonAssigned.value = asset.PersonAssigned || '';
        elements.verifyNotes.value = asset.Notes || '';

        elements.verifyModal.classList.add('active');
    }

    function closeVerifyModal() {
        if (elements.verifyModal) {
            elements.verifyModal.classList.remove('active');
        }
        currentEditingAssetId = null;
    }

    async function saveVerifyChanges() {
        if (!currentEditingAssetId) {
            showStatus('No asset selected for verify', 'error');
            return;
        }

        const id = currentEditingAssetId;
        const payload = {
            AssetID: elements.verifyAssetId.value.trim(),
            SerialNumber: elements.verifySerial.value.trim(),
            OfficeLocation: elements.verifyOfficeLocation.value.trim(),
            Department: elements.verifyDepartment.value.trim(),
            HeadsetType: elements.verifyHeadsetType.value,
            Status: elements.verifyStatus.value,
            PersonAssigned: elements.verifyPersonAssigned.value.trim(),
            Notes: elements.verifyNotes.value.trim()
        };

        if (!payload.AssetID) {
            showStatus('Asset ID cannot be empty', 'warning');
            return;
        }

        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                method: 'POST',
                body: JSON.stringify(payload),
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-Type': 'application/json;odata=nometadata',
                    'X-RequestDigest': requestDigest,
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*'
                }
            });

            if (response.ok) {
                showStatus('Headset details verified and updated', 'success');
                closeVerifyModal();
                await loadData();
                searchHeadset();
            } else {
                throw new Error('Verify save failed');
            }
        } catch (e) {
            console.error('Error saving verification:', e);
            showStatus('Error updating headset details', 'error');
        }
    }
    
    // Assign headset to a person
    async function assignHeadset(id) {
        const assignInput = document.getElementById('assignToName');
        const assignToName = assignInput ? assignInput.value.trim() : '';
        
        if (!assignToName) {
            showStatus("Please enter a name to assign", "warning");
            return;
        }
        
        const asset = masterData.find(a => a.Id === id);
        if (!asset) return;
        
        const timestamp = new Date().toLocaleString();
        const assignmentEntry = `\n[${timestamp}] Assigned to: ${assignToName} - by: User`;
        const updatedNotes = (asset.Notes || '') + assignmentEntry;
        
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                method: "POST",
                body: JSON.stringify({
                    PersonAssigned: assignToName,
                    Status: "Assigned",
                    Notes: updatedNotes
                }),
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "Content-Type": "application/json;odata=nometadata",
                    "X-RequestDigest": requestDigest,
                    "X-HTTP-Method": "MERGE",
                    "IF-MATCH": "*"
                }
            });
            
            if (response.ok) {
                showStatus(`Successfully assigned to ${assignToName}`, "success");
                await loadData();
                searchHeadset();
            } else {
                throw new Error("Assignment failed");
            }
        } catch (error) {
            showStatus("Error assigning headset", "error");
        }
    }
    
    // Add new headset
    async function addNewHeadset() {
        const assetId = elements.newAssetId.value.trim();
        const serial = elements.newSerial.value.trim();
        const officeLocation = elements.newOfficeLocation.value;
        const department = elements.newDepartment.value;
        const headsetType = elements.newHeadsetType.value;
        const personAssigned = elements.newPersonAssigned.value.trim();
        
        if (!assetId) {
            showStatus("Asset ID is required", "warning");
            return;
        }
        
        if (!officeLocation) {
            showStatus("Office Location is required", "warning");
            return;
        }
        
        if (!headsetType) {
            showStatus("Headset Type is required", "warning");
            return;
        }
        
        const existing = masterData.find(a => a.AssetID === assetId);
        if (existing) {
            showStatus("Asset ID already exists!", "error");
            return;
        }
        
        const timestamp = new Date().toLocaleString();
        let notes = `=== HEADSET CREATED ===\nDate: ${timestamp}\nAdded by: User\nInitial Status: Available`;
        
        if (personAssigned) {
            notes += `\n\n[${timestamp}] Assigned to: ${personAssigned} - by: User`;
        }
        
        const payload = {
            AssetID: assetId,
            OfficeLocation: officeLocation,
            Department: department || "",
            SerialNumber: serial,
            Status: personAssigned ? "Assigned" : "Available",
            PersonAssigned: personAssigned || "",
            HeadsetType: headsetType,
            Notes: notes
        };
        
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items`, {
                method: "POST",
                body: JSON.stringify(payload),
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "Content-Type": "application/json;odata=nometadata",
                    "X-RequestDigest": requestDigest
                }
            });
            
            if (response.ok) {
                showStatus(`Headset ${assetId} added successfully!`, "success");
                await loadData();
                clearAddForm();
                elements.searchAssetId.value = '';
                elements.searchSerial.value = '';
                elements.resultSection.classList.remove('active');
            } else {
                throw new Error("Add failed");
            }
        } catch (error) {
            showStatus("Error adding headset", "error");
        }
    }
    
    // Clear add form
    function clearAddForm() {
        elements.newAssetId.value = '';
        elements.newSerial.value = '';
        if (elements.newOfficeLocation) elements.newOfficeLocation.value = '';
        if (elements.newDepartment) elements.newDepartment.value = '';
        if (elements.newHeadsetType) elements.newHeadsetType.value = '';
        elements.newPersonAssigned.value = '';
        elements.addFormSection.classList.add('hidden');
    }
    
    // Admin functions
    function showAdminLogin() {
        if (elements.loginModal) {
            elements.loginModal.classList.add('active');
        }
    }
    
    function checkAdminLogin() {
        const username = elements.loginUsername.value;
        const password = elements.loginPassword.value;
        
        if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
            isAdminLoggedIn = true;
            elements.loginModal.classList.remove('active');
            elements.userView.classList.add('hidden');
            elements.adminView.classList.remove('hidden');
            loadAdminData();
            showStatus("Welcome to Admin Panel", "success");
        } else {
            showStatus("Invalid credentials!", "error");
        }
    }
    
    function logoutAdmin() {
        isAdminLoggedIn = false;
        elements.adminView.classList.add('hidden');
        elements.userView.classList.remove('hidden');
        elements.loginUsername.value = '';
        elements.loginPassword.value = '';
    }
    
    async function loadAdminData() {
        await loadData();
        updateStatistics();
        renderAdminTable();
        updateFilters();
    }
    
    function updateStatistics() {
        if (elements.totalCount) elements.totalCount.textContent = masterData.length;
        if (elements.assignedCount) elements.assignedCount.textContent = masterData.filter(i => i.Status === "Assigned").length;
        if (elements.availableCount) elements.availableCount.textContent = masterData.filter(i => i.Status === "Available").length;
        if (elements.repairCount) elements.repairCount.textContent = masterData.filter(i => i.Status?.includes("Repair")).length;
    }
    
    function renderAdminTable() {
        if (!elements.mainTbody) return;
        
        const aF = elements.fAssetId ? elements.fAssetId.value.toLowerCase() : '';
        const sF = elements.fSerial ? elements.fSerial.value.toLowerCase() : '';
        const stF = elements.fStatus ? elements.fStatus.value : '';
        const tyF = elements.fType ? elements.fType.value : '';
        
        const filtered = masterData.filter(i => 
            (!aF || (i.AssetID || '').toLowerCase().includes(aF)) &&
            (!sF || (i.SerialNumber || '').toLowerCase().includes(sF)) &&
            (!stF || i.Status === stF) &&
            (!tyF || i.HeadsetType === tyF)
        );
        
        elements.mainTbody.innerHTML = filtered.map(i => `
            <tr data-id="${i.Id}">
                <td><input class="inline-input read-only" value="${escapeHtml(i.AssetID)}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.OfficeLocation)}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.Department)}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.SerialNumber)}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.Status)}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.PersonAssigned || '')}" readonly></td>
                <td><input class="inline-input read-only" value="${escapeHtml(i.HeadsetType)}" readonly></td>
                <td><div class="note-preview" title="${escapeHtml(i.Notes || '')}">${escapeHtml((i.Notes || '').substring(0, 50))}${(i.Notes || '').length > 50 ? '...' : ''}</div></td>
                <td>
                    <button class="edit-asset-btn btn-primary" data-id="${i.Id}" style="padding: 4px 8px; font-size: 11px;">Edit</button>
                    <button class="delete-asset-btn btn-danger" data-id="${i.Id}" style="padding: 4px 8px; font-size: 11px; margin-left: 5px;">Delete</button>
                </td>
            </tr>
        `).join('');
        
        // Attach edit/delete event listeners
        document.querySelectorAll('.edit-asset-btn').forEach(btn => {
            btn.addEventListener('click', () => editAsset(parseInt(btn.getAttribute('data-id'))));
        });
        
        document.querySelectorAll('.delete-asset-btn').forEach(btn => {
            btn.addEventListener('click', () => deleteAsset(parseInt(btn.getAttribute('data-id'))));
        });
    }
    
    function updateFilters() {
        if (elements.fStatus) {
            while (elements.fStatus.options.length > 1) elements.fStatus.remove(1);
            choiceOptions.Status.forEach(c => elements.fStatus.add(new Option(c, c)));
        }
        
        if (elements.fType) {
            while (elements.fType.options.length > 1) elements.fType.remove(1);
            choiceOptions.HeadsetType.forEach(c => elements.fType.add(new Option(c, c)));
        }
    }
    
    async function editAsset(id) {
        const asset = masterData.find(a => a.Id === id);
        if (!asset) return;
        
        const newPerson = prompt("Enter new person assignment (leave empty to keep current):", asset.PersonAssigned || "");
        if (newPerson === null) return;
        
        const timestamp = new Date().toLocaleString();
        let notes = asset.Notes || '';
        
        if (newPerson !== asset.PersonAssigned) {
            notes += `\n[${timestamp}] Assignment changed: ${asset.PersonAssigned || 'Unassigned'} → ${newPerson || 'Unassigned'} - by: Admin`;
        }
        
        const status = newPerson ? "Assigned" : "Available";
        
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                method: "POST",
                body: JSON.stringify({
                    PersonAssigned: newPerson || "",
                    Status: status,
                    Notes: notes
                }),
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "Content-Type": "application/json;odata=nometadata",
                    "X-RequestDigest": requestDigest,
                    "X-HTTP-Method": "MERGE",
                    "IF-MATCH": "*"
                }
            });
            
            if (response.ok) {
                showStatus("Asset updated successfully", "success");
                await loadAdminData();
            } else {
                throw new Error("Update failed");
            }
        } catch (error) {
            showStatus("Error updating asset", "error");
        }
    }
    
    async function deleteAsset(id) {
        if (!confirm("Are you sure you want to delete this asset?")) return;
        
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                method: "POST",
                headers: {
                    "X-RequestDigest": requestDigest,
                    "X-HTTP-Method": "DELETE",
                    "IF-MATCH": "*"
                }
            });
            
            if (response.ok) {
                showStatus("Asset deleted successfully", "success");
                await loadAdminData();
            } else {
                throw new Error("Delete failed");
            }
        } catch (error) {
            showStatus("Error deleting asset", "error");
        }
    }
    
    function switchTab(tab) {
        if (tab === 'view') {
            elements.tabView.classList.add('active');
            elements.tabBulk.classList.remove('active');
            elements.viewSection.classList.remove('hidden');
            elements.bulkSection.classList.add('hidden');
        } else {
            elements.tabBulk.classList.add('active');
            elements.tabView.classList.remove('active');
            elements.bulkSection.classList.remove('hidden');
            elements.viewSection.classList.add('hidden');
        }
    }
    
    // CSV Functions
    function downloadTemplate() {
        const csv = `Asset ID,Office Location,Department,Serial Number,Status,Person Assigned,Headset Type,Notes
HT-1001,Main Office,IT,SN123456,Available,John Doe,Logitech Mono USB Headset,New headset
HT-1002,Branch Office,Customer Service,SN789012,Assigned,Jane Smith,Logitech Stereo USB Headset,For remote work`;
        
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = "Headset_Import_Template.csv";
        link.click();
    }
    
    function parseCSVLine(line) {
        const result = [];
        let current = "";
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch === ',' && !inQuotes) {
                result.push(current);
                current = "";
            } else {
                current += ch;
            }
        }
        result.push(current);
        return result.map(v => v.trim().replace(/^"|"$/g, ''));
    }
    
    async function handleCSVFile(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            const text = e.target.result;
            const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
            
            if (lines.length < 2) {
                elements.csvSummary.textContent = "CSV file must have headers and data";
                return;
            }
            
            const headers = parseCSVLine(lines[0]);
            const normalizedHeaders = headers.map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));
            
            csvRows = [];
            
            for (let i = 1; i < lines.length; i++) {
                const cells = parseCSVLine(lines[i]);
                const row = {};
                for (let h = 0; h < headers.length; h++) {
                    if (h < cells.length) {
                        row[normalizedHeaders[h]] = cells[h] || "";
                    }
                }
                
                const assetID = row["assetid"] || "";
                let existingId = null;
                
                if (assetID) {
                    existingId = masterData.find(a => a.AssetID === assetID)?.Id || null;
                }
                
                csvRows.push({
                    index: i,
                    payload: {
                        AssetID: assetID,
                        OfficeLocation: findMatchingChoice('OfficeLocation', row["officelocation"] || ""),
                        Department: findMatchingChoice('Department', row["department"] || ""),
                        SerialNumber: row["serialnumber"] || "",
                        Status: findMatchingChoice('Status', row["status"] || ""),
                        PersonAssigned: row["personassigned"] || "",
                        HeadsetType: findMatchingChoice('HeadsetType', row["headsettype"] || ""),
                        Notes: row["notes"] || ""
                    },
                    existingId: existingId,
                    include: assetID && !existingId
                });
            }
            
            renderCsvPreview();
            elements.startImportBtn.classList.remove('hidden');
            elements.csvSummary.textContent = `Loaded ${csvRows.length} rows. ${csvRows.filter(r => r.include).length} ready for import.`;
        };
        reader.readAsText(file);
    }
    
    function renderCsvPreview() {
        if (!elements.csvPreviewBody) return;
        
        elements.csvPreviewBody.innerHTML = csvRows.map((row, idx) => `
            <tr>
                <td>${row.index}</td>
                <td><input type="checkbox" class="csv-checkbox" ${row.include ? 'checked' : ''} ${!row.payload.AssetID ? 'disabled' : ''} data-index="${idx}"></td>
                <td>${escapeHtml(row.payload.AssetID)}</td>
                <td>${escapeHtml(row.payload.OfficeLocation)}</td>
                <td>${escapeHtml(row.payload.Department)}</td>
                <td>${escapeHtml(row.payload.SerialNumber)}</td>
                <td>${escapeHtml(row.payload.Status)}</td>
                <td>${escapeHtml(row.payload.PersonAssigned)}</td>
                <td>${escapeHtml(row.payload.HeadsetType)}</td>
            </tr>
        `).join('');
        
        document.querySelectorAll('.csv-checkbox').forEach(cb => {
            cb.addEventListener('change', function() {
                const index = parseInt(this.getAttribute('data-index'));
                if (csvRows[index] && csvRows[index].payload.AssetID) {
                    csvRows[index].include = this.checked;
                }
            });
        });
    }
    
    async function importCsvRows() {
        const rowsToImport = csvRows.filter(r => r.include && r.payload.AssetID);
        
        if (rowsToImport.length === 0) {
            showStatus("No rows selected for import", "warning");
            return;
        }
        
        let successCount = 0;
        let errorCount = 0;
        
        for (const row of rowsToImport) {
            try {
                const timestamp = new Date().toLocaleString();
                const notes = `=== HEADSET CREATED ===\nDate: ${timestamp}\nAdded by: Admin (Bulk Import)\nInitial Status: ${row.payload.Status || 'Available'}`;
                
                const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items`, {
                    method: "POST",
                    body: JSON.stringify({
                        ...row.payload,
                        Notes: notes
                    }),
                    headers: {
                        "Accept": "application/json;odata=nometadata",
                        "Content-Type": "application/json;odata=nometadata",
                        "X-RequestDigest": requestDigest
                    }
                });
                
                if (response.ok) {
                    successCount++;
                } else {
                    errorCount++;
                }
            } catch (error) {
                errorCount++;
            }
        }
        
        showStatus(`Import complete: ${successCount} added, ${errorCount} failed`, successCount > 0 ? "success" : "error");
        await loadAdminData();
        elements.csvFile.value = "";
        elements.csvPreviewBody.innerHTML = "";
        elements.startImportBtn.classList.add('hidden');
    }
    
    // Utility Functions
    function showStatus(message, type) {
        elements.statusBox.textContent = message;
        elements.statusBox.className = `status-box ${type}`;
        elements.statusBox.style.display = 'block';
        setTimeout(() => elements.statusBox.style.display = 'none', 5000);
    }
    
    function escapeHtml(text) {
        if (!text) return '';
        return text.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }
    
    // Start the application
    init();
})();