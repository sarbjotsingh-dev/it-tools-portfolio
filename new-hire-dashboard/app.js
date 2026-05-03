(function() {
    const LIST_GUID = "YOUR-RECOGNITION-LIST-GUID";
    const IT_SITE_URL = "https://your-tenant.sharepoint.com/sites/YOUR-SITE-NAME";
    const IT_LIST_GUID = "YOUR-NEWHIRES-LIST-GUID";
    // Prefer current site URL from SharePoint context to avoid cross-site issues in modern pages
    const SITE_URL = (window.COMPANY_SITE_URL) || (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl)
        || (function() {
            try {
                const p = location.pathname.toLowerCase();
                if (p.startsWith('/sites/') || p.startsWith('/teams/')) {
                    const parts = location.pathname.split('/').slice(0, 4).join('/');
                    return location.origin + parts;
                }
                return location.origin;
            } catch (_) {
                return "https://your-tenant.sharepoint.com";
            }
        })();
    
    let listType = ""; 
    let masterData = [];
    let csvRows = [];
    let selectedIds = new Set();
    let requestDigest = "";
    let currentSort = { field: 'hiredate', order: 'desc' };
    let filteredData = [];
    let deleteRecordId = null;
    let allDepartments = new Set();
    let allLocations = new Set();
    let currentCSVBlobUrl = null;
    
    // Track current filter values
    let currentFilters = {
        fName: '',
        fDept: '',
        fLocation: '',
        fHireDate: '',
        fStatus: 'pending'
    };
    
    // Flag to prevent recursive filter updates
    let updatingFilters = false;

    // Location mapping for country codes
    const locationCountryCode = {
        'big newton': '+1',
        'little newton': '+1',
        'north vancouver': '+1',
        'Region-IN': '+91',
        'Region-IN2': '+91',
        'Region-PH2': '+63',
        'legazpi': '+63',
        'Region-PH3': '+63',
        'Region-SA': '+27',
        'capetown': '+27'
    };

    // Enhanced phone number formatter
    function formatPhoneNumber(phone, location) {
        if (!phone) return "";
        
        let cleaned = phone.replace(/[^\d+]/g, '');
        
        if (cleaned.match(/^\+\d{1,3}\s*\d{4,}$/)) {
            return cleaned;
        }
        
        let countryCode = "+1";
        if (location) {
            const locLower = location.toLowerCase();
            for (const [key, code] of Object.entries(locationCountryCode)) {
                if (locLower.includes(key)) {
                    countryCode = code;
                    break;
                }
            }
        }
        
        const digitsOnly = cleaned.replace(/\D/g, '');
        
        if (countryCode === "+1") {
            if (digitsOnly.length === 10) {
                return `${countryCode} ${digitsOnly.substring(0,3)} ${digitsOnly.substring(3,6)} ${digitsOnly.substring(6)}`;
            } else if (digitsOnly.length === 11 && digitsOnly.startsWith('1')) {
                return `+1 ${digitsOnly.substring(1,4)} ${digitsOnly.substring(4,7)} ${digitsOnly.substring(7)}`;
            }
        } else if (countryCode === "+63") {
            if (digitsOnly.length === 10) {
                return `${countryCode} ${digitsOnly.substring(0,3)} ${digitsOnly.substring(3,6)} ${digitsOnly.substring(6)}`;
            }
        } else if (countryCode === "+91") {
            if (digitsOnly.length === 10) {
                return `${countryCode} ${digitsOnly.substring(0,5)} ${digitsOnly.substring(5)}`;
            }
        } else if (countryCode === "+27") {
            if (digitsOnly.length === 9) {
                return `${countryCode} ${digitsOnly.substring(0,2)} ${digitsOnly.substring(2,5)} ${digitsOnly.substring(5)}`;
            }
        }
        
        return `${countryCode} ${digitsOnly}`;
    }

    function validatePhoneForCountry(digitsOnly, countryCode, rawPhone) {
        if (!digitsOnly) return { valid: false, message: "Phone number is empty" };
        if (countryCode === "+1") {
            if (digitsOnly.length === 10) {
                if (digitsOnly[0] === '0' || digitsOnly[0] === '1') {
                    return { valid: false, message: `US/Region-CA area code cannot start with "${digitsOnly[0]}" — got "${rawPhone}". Enter 10 digits (no country code) or 11 digits starting with 1` };
                }
                return { valid: true };
            } else if (digitsOnly.length === 11 && digitsOnly[0] === '1') {
                if (digitsOnly[1] === '0' || digitsOnly[1] === '1') {
                    return { valid: false, message: `US/Region-CA area code cannot start with "${digitsOnly[1]}" — got "${rawPhone}"` };
                }
                return { valid: true };
            }
            return { valid: false, message: `US/Region-CA numbers must be 10 digits or 11 digits with country code 1 — got ${digitsOnly.length} digits for "${rawPhone}"` };
        } else if (countryCode === "+91") {
            if (digitsOnly.length === 10 && ['6','7','8','9'].includes(digitsOnly[0])) return { valid: true };
            return { valid: false, message: `Region-IN numbers must be 10 digits starting with 6–9 — got "${rawPhone}"` };
        } else if (countryCode === "+63") {
            if (digitsOnly.length === 10 && digitsOnly[0] === '9') return { valid: true }; // +63 stripped: 9XXXXXXXXX
            if (digitsOnly.length === 11 && digitsOnly[0] === '0') return { valid: true }; // local: 09XXXXXXXXX
            return { valid: false, message: `Region-PH numbers must be 10 digits starting with 9 (international, e.g. +63 9XX XXX XXXX) or 11 digits starting with 0 (local, e.g. 09XX XXX XXXX) — got "${rawPhone}"` };
        } else if (countryCode === "+27") {
            if (digitsOnly.length === 9) return { valid: true };
            if (digitsOnly.length === 10 && digitsOnly[0] === '0') return { valid: true };
            return { valid: false, message: `South Africa numbers must be 9 digits or 10 digits starting with 0 — got "${rawPhone}"` };
        }
        return { valid: true };
    }

    function extractNames(fullName) {
        const nameParts = fullName.trim().split(/\s+/);
        let firstName = "";
        let lastName = "";
        
        if (nameParts.length === 1) {
            firstName = nameParts[0];
            lastName = "";
        } else if (nameParts.length === 2) {
            firstName = nameParts[0];
            lastName = nameParts[1];
        } else {
            firstName = nameParts[0];
            lastName = nameParts.slice(1).join(" ");
        }
        
        return { firstName, lastName };
    }

    function generateEmailUsername(firstName, lastName) {
        const first = (firstName || "").toLowerCase().replace(/[^a-z]/g, '');
        const last = (lastName || "").toLowerCase().replace(/[^a-z]/g, '');
        if (!first && !last) return "";
        return `${first}.${last}@company.com`;
    }

    function generateIloanId(firstName, lastName) {
        const firstPart = (firstName || "").substring(0, 4).toUpperCase();
        const lastPart = (lastName || "").substring(0, 5).toUpperCase();
        return `${firstPart}.${lastPart}`;
    }

    function isEmailRequired(location) {
        if (!location) return false;
        const locLower = location.toLowerCase();
        return locLower.includes("Region-PH2") || 
               locLower.includes("legazpi") || 
               locLower.includes("Region-PH3") || 
               locLower.includes("Region-SA") || 
               locLower.includes("capetown") || 
               locLower.includes("Region-IN2");
    }

    function formatDateForDisplay(dateString) {
        if (!dateString) return "";
        try {
            const date = new Date(dateString);
            if (isNaN(date.getTime())) return dateString;
            
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const day = date.getDate().toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${month}/${day}/${year}`;
        } catch (e) {
            return dateString;
        }
    }

    function formatDateForSharePoint(dateString) {
        if (!dateString) return null;
        
        try {
            let date;
            
            if (dateString.includes('/')) {
                const parts = dateString.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1;
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    date = new Date(year, month, day);
                }
            } else {
                date = new Date(dateString);
            }
            
            if (!date || isNaN(date.getTime())) {
                return null;
            }
            
            return date.toISOString();
        } catch (e) {
            return null;
        }
    }

    async function init() {
        try {
            bindEvents();
        } catch (e) {
            console.error('Event binding failed:', e);
        }
        await fetchType();
        await refreshDigest();
        await loadData();
        setupSorting();
    }

    // In modern pages, ensure DOM is ready and elements exist before init
    function whenDomReady(cb) {
        if (document.readyState === 'complete' || document.readyState === 'interactive') {
            cb();
        } else {
            document.addEventListener('DOMContentLoaded', cb, { once: true });
        }
    }

    function waitForElements(selectors, timeoutMs = 8000) {
        return new Promise(resolve => {
            const start = Date.now();
            (function check() {
                const missing = selectors.some(sel => !document.querySelector(sel));
                if (!missing) return resolve(true);
                if (Date.now() - start > timeoutMs) return resolve(false);
                setTimeout(check, 100);
            })();
        });
    }

    function bindEvents() {
        const safeBind = (id, evt, handler) => {
            const el = document.getElementById(id);
            if (el) el[evt] = handler;
        };

        safeBind('tabView', 'onclick', () => switchTab('view'));
        safeBind('tabBulk', 'onclick', () => switchTab('bulk'));
        
        safeBind('btnRefresh', 'onclick', refreshData);
        safeBind('btnManualLoad', 'onclick', manualLoadData);
        safeBind('btnClearFilters', 'onclick', clearAllFilters);
        
        safeBind('templateCSCSV', 'onclick', () => downloadTemplate('customer-service'));
        safeBind('templateAssessmentCSV', 'onclick', () => downloadTemplate('assessment'));
        safeBind('templateCollectionsCSV', 'onclick', () => downloadTemplate('collections'));
        
        safeBind('csvFile', 'onchange', handleCSV);
        safeBind('btnStart', 'onclick', importCsvRows);
        safeBind('btnDeleteSelected', 'onclick', deleteSelected);
        safeBind('btnMarkCompleted', 'onclick', markAsCompleted);
        safeBind('btnEmailSelected', 'onclick', showEmailModal);
        safeBind('btnCreateFolders', 'onclick', markCreateFolderComplete);
        safeBind('btnCreateBookings', 'onclick', markBookingsProfileComplete);
        safeBind('btnPushToIT', 'onclick', pushToITList);
        safeBind('btnLMProfileCSV', 'onclick', downloadLMProfileCSV);
        safeBind('selectAll', 'onclick', toggleAllCheckboxes);
        
        safeBind('btnSelectAllValid', 'onclick', selectAllValidCsv);
        safeBind('btnDeselectAll', 'onclick', deselectAllCsv);
        safeBind('selectAllCsv', 'onclick', toggleAllCsv);
        
        safeBind('confirmationCancel', 'onclick', hideConfirmation);
        safeBind('confirmationConfirm', 'onclick', performDeleteRecord);
        safeBind('confirmationOverlay', 'onclick', hideConfirmation);
        
        safeBind('emailCancel', 'onclick', hideEmailModal);
        safeBind('emailOpen', 'onclick', openEmailClient);
        safeBind('emailOverlay', 'onclick', hideEmailModal);

        safeBind('dupAlertOk', 'onclick', hideDuplicateAlert);
        safeBind('dupAlertOverlay', 'onclick', hideDuplicateAlert);
        safeBind('emailTo', 'oninput', updateEmailPreview);
        safeBind('emailSubject', 'oninput', updateEmailPreview);
        
        // Dynamic filter bindings with real-time updates
        const fNameEl = document.getElementById('fName');
        const fDeptEl = document.getElementById('fDept');
        const fLocationEl = document.getElementById('fLocation');
        const fHireDateEl = document.getElementById('fHireDate');
        const fStatusEl = document.getElementById('fStatus');
        
        if (fNameEl) {
            fNameEl.oninput = (e) => {
                currentFilters.fName = e.target.value.toLowerCase();
                updateDynamicFilters();
            };
        }
        
        if (fDeptEl) {
            fDeptEl.onchange = (e) => {
                currentFilters.fDept = e.target.value;
                updateDynamicFilters();
            };
        }
        
        if (fLocationEl) {
            fLocationEl.onchange = (e) => {
                currentFilters.fLocation = e.target.value;
                updateDynamicFilters();
            };
        }
        
        if (fHireDateEl) {
            fHireDateEl.onchange = (e) => {
                currentFilters.fHireDate = e.target.value;
                updateDynamicFilters();
            };
        }
        
        if (fStatusEl) {
            fStatusEl.onchange = (e) => {
                currentFilters.fStatus = e.target.value;
                updateDynamicFilters();
            };
        }
    }

    // Clear all filters function
    function clearAllFilters() {
        // Reset filter values in tracking object
        currentFilters = {
            fName: '',
            fDept: '',
            fLocation: '',
            fHireDate: '',
            fStatus: 'pending'
        };
        
        // Reset UI elements
        const fNameEl = document.getElementById('fName');
        const fDeptEl = document.getElementById('fDept');
        const fLocationEl = document.getElementById('fLocation');
        const fHireDateEl = document.getElementById('fHireDate');
        const fStatusEl = document.getElementById('fStatus');
        
        if (fNameEl) fNameEl.value = '';
        if (fDeptEl) fDeptEl.value = '';
        if (fLocationEl) fLocationEl.value = '';
        if (fHireDateEl) fHireDateEl.value = '';
        if (fStatusEl) fStatusEl.value = 'pending';
        
        // Clear selections
        selectedIds.clear();
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
        
        const selectAll = document.getElementById('selectAll');
        if (selectAll) selectAll.checked = false;
        
        // Update filters to show all data
        updateDynamicFilters();
        
        showStatus('All filters cleared', 'info');
    }

    // Main function to update filters dynamically
    function updateDynamicFilters() {
        if (updatingFilters) return;
        updatingFilters = true;
        
        try {
            // Apply current filters to get filtered dataset
            const filteredDataset = applyFiltersToData(masterData, currentFilters);
            
            // Update filter dropdowns based on filtered dataset
            updateDepartmentFilter(filteredDataset);
            updateLocationFilter(filteredDataset);
            updateHireDateFilter(filteredDataset);
            
            // Update the main display
            renderMain();
            
            // Update selection buttons
            updateBulkDeleteButton();
            updateEmailButton();
            updateMarkCompletedButton();
            updateFolderButtons();
            
        } finally {
            updatingFilters = false;
        }
    }

    // Apply filters to a dataset
    function applyFiltersToData(data, filters) {
        return data.filter(item => {
            // Name filter
            let mName = true;
            if (filters.fName) {
                mName = String(item.EmployeeName || item.FirstName || "").toLowerCase().includes(filters.fName);
            }
            
            // Department filter
            let mDept = true;
            if (filters.fDept) {
                mDept = item.Department === filters.fDept;
            }
            
            // Location filter
            let mLocation = true;
            if (filters.fLocation) {
                mLocation = item.Location === filters.fLocation;
            }
            
            // Hire date filter
            let mHireDate = true;
            if (filters.fHireDate) {
                const formattedHireDate = item.HireDate ? formatDateForDisplay(item.HireDate) : '';
                mHireDate = formattedHireDate === filters.fHireDate;
            }
            
            // Status filter
            let mStatus = true;
            if (filters.fStatus === 'pending') {
                mStatus = item.StatusOfNewHire === 'Pending';
            } else if (filters.fStatus === 'completed') {
                mStatus = item.StatusOfNewHire === 'Completed';
            } else if (filters.fStatus === 'all') {
                mStatus = true;
            }
            
            return mName && mDept && mLocation && mHireDate && mStatus;
        });
    }

    // Update department filter dropdown based on filtered data
    function updateDepartmentFilter(filteredDataset) {
        const deptSelect = document.getElementById('fDept');
        if (!deptSelect) return;
        
        // Get unique departments from filtered dataset
        const departments = new Set();
        filteredDataset.forEach(item => {
            if (item.Department && item.Department.trim() !== '') {
                departments.add(item.Department);
            }
        });
        
        // Sort departments alphabetically
        const sortedDepts = Array.from(departments).sort();
        
        // Get currently selected value
        const currentValue = deptSelect.value;
        
        // Clear existing options (keep first "All" option)
        while (deptSelect.options.length > 1) deptSelect.remove(1);
        
        // Add new options
        sortedDepts.forEach(dept => {
            const option = document.createElement('option');
            option.value = dept;
            option.textContent = dept;
            deptSelect.appendChild(option);
        });
        
        // Restore selected value if it still exists
        if (currentValue && departments.has(currentValue)) {
            deptSelect.value = currentValue;
            currentFilters.fDept = currentValue;
        } else if (deptSelect.value === '') {
            currentFilters.fDept = '';
        } else {
            // Current selection no longer available, reset it
            deptSelect.value = '';
            currentFilters.fDept = '';
        }
    }

    // Update location filter dropdown based on filtered data
    function updateLocationFilter(filteredDataset) {
        const locationSelect = document.getElementById('fLocation');
        if (!locationSelect) return;
        
        // Get unique locations from filtered dataset
        const locations = new Set();
        filteredDataset.forEach(item => {
            if (item.Location && item.Location.trim() !== '') {
                locations.add(item.Location);
            }
        });
        
        // Sort locations alphabetically
        const sortedLocs = Array.from(locations).sort();
        
        // Get currently selected value
        const currentValue = locationSelect.value;
        
        // Clear existing options (keep first "All" option)
        while (locationSelect.options.length > 1) locationSelect.remove(1);
        
        // Add new options
        sortedLocs.forEach(location => {
            const option = document.createElement('option');
            option.value = location;
            option.textContent = location;
            locationSelect.appendChild(option);
        });
        
        // Restore selected value if it still exists
        if (currentValue && locations.has(currentValue)) {
            locationSelect.value = currentValue;
            currentFilters.fLocation = currentValue;
        } else if (locationSelect.value === '') {
            currentFilters.fLocation = '';
        } else {
            // Current selection no longer available, reset it
            locationSelect.value = '';
            currentFilters.fLocation = '';
        }
    }

    // Update hire date filter dropdown based on filtered data
    function updateHireDateFilter(filteredDataset) {
        const hireDateSelect = document.getElementById('fHireDate');
        if (!hireDateSelect) return;
        
        // Get unique hire dates from filtered dataset
        const hireDates = new Set();
        filteredDataset.forEach(item => {
            if (item.HireDate && (item.StatusOfNewHire === 'Pending' || item.StatusOfNewHire === 'Completed')) {
                const formattedDate = formatDateForDisplay(item.HireDate);
                hireDates.add(formattedDate);
            }
        });
        
        // Sort hire dates chronologically
        const sortedDates = Array.from(hireDates).sort((a, b) => {
            const dateA = new Date(a);
            const dateB = new Date(b);
            return dateA - dateB;
        });
        
        // Get currently selected value
        const currentValue = hireDateSelect.value;
        
        // Clear existing options (keep first "All" option)
        while (hireDateSelect.options.length > 1) hireDateSelect.remove(1);
        
        // Add new options
        sortedDates.forEach(date => {
            const option = document.createElement('option');
            option.value = date;
            option.textContent = date;
            hireDateSelect.appendChild(option);
        });
        
        // Restore selected value if it still exists
        if (currentValue && hireDates.has(currentValue)) {
            hireDateSelect.value = currentValue;
            currentFilters.fHireDate = currentValue;
        } else if (hireDateSelect.value === '') {
            currentFilters.fHireDate = '';
        } else {
            // Current selection no longer available, reset it
            hireDateSelect.value = '';
            currentFilters.fHireDate = '';
        }
    }

    function clearSelectionsOnFilter() {
        selectedIds.clear();
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
        const selectAll = document.getElementById('selectAll');
        if (selectAll) selectAll.checked = false;
    }

    async function refreshData() {
        showStatus('Refreshing data...', 'info');
        await loadData();
        showStatus('Data refreshed successfully', 'success');
    }

    async function manualLoadData() {
        showStatus('Manually loading data...', 'info');
        await loadData();
        showStatus('Data loaded successfully', 'success');
    }

    function selectAllValidCsv() {
        csvRows.forEach((row) => {
            if (row.isValid && !row.existingId) {
                row.include = true;
            }
        });
        renderCsvPreview();
        updateCsvSelectAllCheckbox();
    }

    function deselectAllCsv() {
        csvRows.forEach(row => {
            row.include = false;
        });
        renderCsvPreview();
        updateCsvSelectAllCheckbox();
    }

    function toggleAllCsv(e) {
        const checked = e.target.checked;
        csvRows.forEach((row) => {
            if (row.isValid) {
                row.include = checked;
            }
        });
        renderCsvPreview();
    }

    function updateCsvSelectAllCheckbox() {
        const selectAllCheckbox = document.getElementById('selectAllCsv');
        if (!selectAllCheckbox) return;
        
        const validRows = csvRows.filter(r => r.isValid);
        const includedValid = validRows.filter(r => r.include);
        
        if (validRows.length === 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        } else if (includedValid.length === validRows.length) {
            selectAllCheckbox.checked = true;
            selectAllCheckbox.indeterminate = false;
        } else if (includedValid.length === 0) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        } else {
            selectAllCheckbox.indeterminate = true;
        }
    }

    function downloadTemplate(templateType) {
        const today = new Date().toISOString().split('T')[0];
        let csvContent = '';
        let filename = '';
        
        if (templateType === 'customer-service') {
            csvContent = `Name,Location,Position,Department,Phone Number,External Email,Hire Date,Manager
John Doe,Big Newton,Customer Service Representative,Customer Service,11234567890,john.doe@company.com,${today},manager@company.com
Jane Smith,North Vancouver,Customer Service Representative,Customer Service,11234567890,jane.smith@company.com,${today},manager@company.com
Mike Johnson,Region-PH2,Tier 1 CSR,Customer Service,11234567890,mike.johnson@company.com,${today},manager@company.com`;
            filename = `Customer_Service_Template_${today}.csv`;
        } else if (templateType === 'assessment') {
            csvContent = `Name,Location,Position,Department,Phone Number,External Email,Hire Date,Manager
Sarah Williams,North Vancouver,Assessor,Assessment,11234567890,sarah.williams@company.com,${today},manager@company.com
David Brown,Region-SA,Assessment Specialist,Assessment,11234567890,david.brown@company.com,${today},manager@company.com
Lisa Davis,Legazpi,Quality Assessor,Assessment,11234567890,lisa.davis@company.com,${today},manager@company.com`;
            filename = `Assessment_Template_${today}.csv`;
        } else if (templateType === 'collections') {
            csvContent = `Name,Location,Position,Department,Phone Number,External Email,Hire Date,Manager
Robert Johnson,Office Location,Account Manager,Collections,11234567890,robert.johnson@company.com,${today},manager@company.com
Maria Garcia,Region-PH3,Collections Specialist,Collections,11234567890,maria.garcia@company.com,${today},manager@company.com
James Wilson,Region-IN2,Collections Agent,Collections,11234567890,james.wilson@company.com,${today},manager@company.com`;
            filename = `Collections_Template_${today}.csv`;
        }
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        
        if (navigator.msSaveBlob) {
            navigator.msSaveBlob(blob, filename);
        } else {
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            link.style.display = 'none';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
        
        showStatus(`${templateType.charAt(0).toUpperCase() + templateType.slice(1)} template downloaded successfully`, "success");
    }

    function populateFilterOptions() {
        const filteredDataset = applyFiltersToData(masterData, currentFilters);
        updateDepartmentFilter(filteredDataset);
        updateLocationFilter(filteredDataset);
        updateHireDateFilter(filteredDataset);
        
        const fStatusEl = document.getElementById('fStatus');
        if (fStatusEl && currentFilters.fStatus) {
            fStatusEl.value = currentFilters.fStatus;
        }
    }

    function getUniqueHireDates() {
        const dates = new Set();
        masterData.forEach(item => {
            if (item.HireDate && (item.StatusOfNewHire === 'Pending' || item.StatusOfNewHire === 'Completed')) {
                const formattedDate = formatDateForDisplay(item.HireDate);
                dates.add(formattedDate);
            }
        });
        return Array.from(dates).sort((a, b) => {
            const dateA = new Date(a);
            const dateB = new Date(b);
            return dateA - dateB;
        });
    }

    function showConfirmation(recordId) {
        deleteRecordId = recordId;
        document.getElementById('confirmationOverlay').classList.remove('hidden');
        document.getElementById('confirmationBox').classList.remove('hidden');
    }

    function hideConfirmation() {
        deleteRecordId = null;
        document.getElementById('confirmationOverlay').classList.add('hidden');
        document.getElementById('confirmationBox').classList.add('hidden');
    }

    async function performDeleteRecord() {
        if (!deleteRecordId) {
            hideConfirmation();
            return;
        }
        
        const id = deleteRecordId;
        hideConfirmation();
        
        if (!requestDigest) await refreshDigest();
        
        try {
            const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                method: "POST",
                headers: { 
                    "X-RequestDigest": requestDigest, 
                    "X-HTTP-Method": "DELETE", 
                    "IF-MATCH": "*" 
                },
                credentials: 'same-origin'
            });
            
            if (response.ok) {
                showStatus("Record deleted successfully", "success");
                await refreshDataPreserveState();
            } else {
                showStatus("Error deleting record", "error");
            }
        } catch (error) {
            showStatus("Error deleting record", "error");
        }
    }

    async function markAsCompleted() {
        if (selectedIds.size === 0) {
            showStatus("Please select records to mark as completed", "warning");
            return;
        }
        
        if (!requestDigest) await refreshDigest();
        
        let updated = 0;
        let failed = 0;
        
        for(let id of selectedIds) {
            try {
                const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                    method: "POST",
                    headers: {
                        "Accept": "application/json;odata=nometadata",
                        "Content-Type": "application/json;odata=nometadata",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": requestDigest,
                        "IF-MATCH": "*"
                    },
                    credentials: 'same-origin',
                    body: JSON.stringify({ StatusOfNewHire: "Completed" })
                });

                if (resp.ok) {
                    const record = masterData.find(item => item.Id === id);
                    if (record) record.StatusOfNewHire = "Completed";
                    updated++;
                } else {
                    failed++;
                }
            } catch (error) {
                failed++;
            }
        }
        
        selectedIds.clear();
        showStatus(`Successfully marked ${updated} records as completed${failed > 0 ? `, ${failed} failed` : ''}`, failed === 0 ? "success" : "warning");
        
        updateDynamicFilters();
    }

    async function markCreateFolderComplete() {
        if (selectedIds.size === 0) {
            showStatus("Please select records to mark Folder as created", "warning");
            return;
        }
        
        const recordsToUpdate = [];
        for(let id of selectedIds) {
            const record = masterData.find(item => item.Id === id);
            if (record && record.CreateFolder !== "Yes") {
                recordsToUpdate.push({ id, record });
            }
        }
        
        if (recordsToUpdate.length === 0) {
            showStatus("Selected records already have Folders created", "info");
            return;
        }
        
        showStatus(`Creating Folders for ${recordsToUpdate.length} records...`, "info");
        
        if (!requestDigest) await refreshDigest();
        
        let updated = 0;
        let failed = 0;
        
        for (const { id, record } of recordsToUpdate) {
            try {
                const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                    method: "POST",
                    headers: {
                        "Accept": "application/json;odata=nometadata",
                        "Content-Type": "application/json;odata=nometadata",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": requestDigest,
                        "IF-MATCH": "*"
                    },
                    credentials: 'same-origin',
                    body: JSON.stringify({ CreateFolder: "Yes" })
                });

                if (resp.ok) {
                    record.CreateFolder = "Yes";
                    updated++;
                } else {
                    failed++;
                }
            } catch (error) {
                failed++;
                console.error(`Error creating folder for ID ${id}:`, error);
            }
        }
        
        showStatus(`Successfully created folders for ${updated} records${failed > 0 ? `, ${failed} failed` : ''}`, failed === 0 ? "success" : "warning");
        
        updateDynamicFilters();
        
        selectedIds.clear();
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
    }

    async function markBookingsProfileComplete() {
        if (selectedIds.size === 0) {
            showStatus("Please select records to mark Bookings Profile as created", "warning");
            return;
        }
        
        const recordsToUpdate = [];
        for(let id of selectedIds) {
            const record = masterData.find(item => item.Id === id);
            if (record && record.CompanyBookings !== "Yes") {
                recordsToUpdate.push({ id, record });
            }
        }
        
        if (recordsToUpdate.length === 0) {
            showStatus("Selected records already have Bookings Profile created", "info");
            return;
        }
        
        showStatus(`Updating Bookings Profile for ${recordsToUpdate.length} records...`, "info");
        
        if (!requestDigest) await refreshDigest();
        
        let updated = 0;
        let failed = 0;
        
        for (const { id, record } of recordsToUpdate) {
            try {
                const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                    method: "POST",
                    headers: {
                        "Accept": "application/json;odata=nometadata",
                        "Content-Type": "application/json;odata=nometadata",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": requestDigest,
                        "IF-MATCH": "*"
                    },
                    credentials: 'same-origin',
                    body: JSON.stringify({ CompanyBookings: "Yes" })
                });

                if (resp.ok) {
                    record.CompanyBookings = "Yes";
                    updated++;
                } else {
                    failed++;
                }
            } catch (error) {
                failed++;
                console.error(`Error updating Bookings Profile for ID ${id}:`, error);
            }
        }
        
        showStatus(`Successfully marked ${updated} records as Bookings Profile created${failed > 0 ? `, ${failed} failed` : ''}`, failed === 0 ? "success" : "warning");
        
        updateDynamicFilters();
        
        selectedIds.clear();
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
    }

    function showEmailModal() {
        if (selectedIds.size === 0) {
            showStatus("Please select records to email", "warning");
            return;
        }
        
        updateEmailPreview();
        document.getElementById('emailOverlay').classList.remove('hidden');
        document.getElementById('emailModal').classList.remove('hidden');
    }

    function hideEmailModal() {
        document.getElementById('emailOverlay').classList.add('hidden');
        document.getElementById('emailModal').classList.add('hidden');
    }

    function showDuplicateAlert(duplicateRows) {
        const rows = duplicateRows.map(r => {
            const existing = masterData.find(m => m.Id === r.existingId);
            if (!existing) return '';
            return `<tr>
                <td style="padding:7px 10px;border-bottom:1px solid #edebe9;">${escapeHtml(existing.EmployeeName || '')}</td>
                <td style="padding:7px 10px;border-bottom:1px solid #edebe9;">${escapeHtml(existing.Department || '')}</td>
                <td style="padding:7px 10px;border-bottom:1px solid #edebe9;">${escapeHtml(formatDateForDisplay(existing.HireDate) || '')}</td>
                <td style="padding:7px 10px;border-bottom:1px solid #edebe9;">${escapeHtml(existing.PhoneNumber || '')}</td>
            </tr>`;
        }).join('');
        document.getElementById('dupAlertTable').innerHTML = `
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <thead>
                    <tr style="background:#f3f2f1;">
                        <th style="padding:7px 10px;text-align:left;border-bottom:2px solid #edebe9;font-weight:600;">Name</th>
                        <th style="padding:7px 10px;text-align:left;border-bottom:2px solid #edebe9;font-weight:600;">Department</th>
                        <th style="padding:7px 10px;text-align:left;border-bottom:2px solid #edebe9;font-weight:600;">Hire Date</th>
                        <th style="padding:7px 10px;text-align:left;border-bottom:2px solid #edebe9;font-weight:600;">Phone</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>`;
        document.getElementById('dupAlertOverlay').classList.remove('hidden');
        document.getElementById('dupAlertModal').classList.remove('hidden');
    }

    function hideDuplicateAlert() {
        document.getElementById('dupAlertOverlay').classList.add('hidden');
        document.getElementById('dupAlertModal').classList.add('hidden');
    }

    function escapeCSVValue(value) {
        const stringValue = value == null ? '' : String(value);
        return `"${stringValue.replace(/"/g, '""')}"`;
    }

    function generateCSVForEmail() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        
        if (selectedRecords.length === 0) return "";
        
        const headers = [
            'Full name', 'First Name', 'Last Name', 'M365 Username', 'M365 Password',
            'Five9 Username', 'Five9 Password', 'Five9 Station ID', 'ILOAN ID',
            'Department', 'Position', 'Hire Date', 'Location', 'Manager',
            'Phone Number', 'External Email'
        ];
        
        let csvContent = headers.join(',') + '\n';
        
        selectedRecords.forEach(record => {
            const row = [
                escapeCSVValue(record.EmployeeName),
                escapeCSVValue(record.FirstName),
                escapeCSVValue(record.LastName),
                escapeCSVValue(record.M365Username_x002f_Computerusern),
                escapeCSVValue(record.Windows_x002f_M365Password),
                escapeCSVValue(record.Five9Username),
                escapeCSVValue(record.Five9Password),
                escapeCSVValue(record.Five9StationID),
                escapeCSVValue(record.ILOANID),
                escapeCSVValue(record.Department),
                escapeCSVValue(record.Position),
                escapeCSVValue(formatDateForDisplay(record.HireDate)),
                escapeCSVValue(record.Location),
                escapeCSVValue(record.Manager),
                escapeCSVValue(record.PhoneNumber),
                escapeCSVValue(record.ExternalEmail)
            ];
            csvContent += row.join(',') + '\n';
        });
        
        return csvContent;
    }

    function createCSVDownloadLink() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        
        if (selectedRecords.length === 0) return "";
        
        if (currentCSVBlobUrl) {
            URL.revokeObjectURL(currentCSVBlobUrl);
        }
        
        const csvContent = generateCSVForEmail();
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        currentCSVBlobUrl = URL.createObjectURL(blob);
        
        const date = new Date().toISOString().split('T')[0];
        const filename = `New_Hires_${date}_${selectedRecords.length}_records.csv`;
        
        return `<a href="${currentCSVBlobUrl}" download="${filename}" class="csv-download-link">📥 Download CSV (${selectedRecords.length} records)</a>`;
    }

    function updateEmailPreview() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        
        if (selectedRecords.length === 0) {
            document.getElementById('emailBodyPreview').textContent = "No records selected.";
            document.getElementById('csvDownloadContainer').innerHTML = "";
            return;
        }
        
        const firstRecord = selectedRecords[0];
        const department = firstRecord.Department || "Various";
        const location = firstRecord.Location || "Various";
        
        let earliestHireDate = null;
        selectedRecords.forEach(record => {
            if (record.HireDate) {
                const hireDate = new Date(record.HireDate);
                if (!earliestHireDate || hireDate < earliestHireDate) {
                    earliestHireDate = hireDate;
                }
            }
        });
        
        const hireDateStr = earliestHireDate ? formatDateForDisplay(earliestHireDate.toISOString()) : "Various";
        
        document.getElementById('emailSubject').value = `New Hire IT Setup - ${department} - ${location} - ${hireDateStr}`;
        
        let emailBody = `Hi IT Team,\n\n`;
        emailBody += `The New Hire file has been updated with ${selectedRecords.length} new employees, starting on ${hireDateStr} at ${location}.\n\n`;
        emailBody += `Please review the attached file for full details.\n\n`;
        emailBody += `Thank you.\n\n`;
        emailBody += `---\n`;
        emailBody += `Summary:\n`;
        emailBody += `• Department: ${department}\n`;
        emailBody += `• Location: ${location}\n`;
        emailBody += `• Start Date: ${hireDateStr}\n\n`;
        
        document.getElementById('emailBodyPreview').textContent = emailBody;
        
        const csvLink = createCSVDownloadLink();
        document.getElementById('csvDownloadContainer').innerHTML = `
            <div style="margin-bottom: 10px;">${csvLink}</div>
            <div style="font-size: 11px; color: #605e5c;">
                Download and attach this CSV to your email (${selectedRecords.length} records)
            </div>
        `;
    }

    function openEmailClient() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        if (selectedRecords.length === 0) return;
        
        const to = document.getElementById('emailTo').value.trim();
        const subject = encodeURIComponent(document.getElementById('emailSubject').value.trim());
        
        const firstRecord = selectedRecords[0];
        const department = firstRecord.Department || "Various";
        const location = firstRecord.Location || "Various";
        
        let earliestHireDate = null;
        selectedRecords.forEach(record => {
            if (record.HireDate) {
                const hireDate = new Date(record.HireDate);
                if (!earliestHireDate || hireDate < earliestHireDate) {
                    earliestHireDate = hireDate;
                }
            }
        });
        
        const hireDateStr = earliestHireDate ? formatDateForDisplay(earliestHireDate.toISOString()) : "Various";
        
        const encDept = encodeURIComponent(department);
        const encLoc = encodeURIComponent(location);
        const encDate = encodeURIComponent(hireDateStr);
        let emailBody = `Hi IT Team,%0A%0A`;
        emailBody += `The New Hire file has been updated with ${selectedRecords.length} new employees, starting on ${encDate} at ${encLoc}.%0A%0A`;
        emailBody += `Please review the attached file for full details.%0A%0A`;
        emailBody += `Thank you.%0A%0A`;
        emailBody += `---%0A`;
        emailBody += `Summary:%0A`;
        emailBody += `%E2%80%A2 Department: ${encDept}%0A`;
        emailBody += `%E2%80%A2 Location: ${encLoc}%0A`;
        emailBody += `%E2%80%A2 Start Date: ${encDate}`;
        
        const mailtoLink = `mailto:${to}?subject=${subject}&body=${emailBody}`;
        window.location.href = mailtoLink;
        
        hideEmailModal();
        showStatus(`Email opened with ${selectedRecords.length} records`, "success");
    }

    function setupSorting() {
        document.querySelectorAll('.sortable').forEach(th => {
            th.addEventListener('click', () => {
                const field = th.getAttribute('data-sort');
                const currentOrder = th.getAttribute('data-order');
                
                if (currentSort.field === field) {
                    currentSort.order = currentOrder === 'asc' ? 'desc' : 'asc';
                } else {
                    currentSort.field = field;
                    currentSort.order = 'asc';
                }
                
                renderMain();
            });
        });
    }

    function sortData(data, field, order) {
        return [...data].sort((a, b) => {
            let aVal = a[field] || '';
            let bVal = b[field] || '';
            
            if (field === 'HireDate') {
                aVal = new Date(aVal).getTime() || 0;
                bVal = new Date(bVal).getTime() || 0;
                return order === 'asc' ? aVal - bVal : bVal - aVal;
            }
            
            aVal = String(aVal).toLowerCase();
            bVal = String(bVal).toLowerCase();
            
            return order === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        });
    }

    function updateSortUI() {
        document.querySelectorAll('.sortable').forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
        });
        
        const currentTh = document.querySelector(`.sortable[data-sort="${currentSort.field.toLowerCase()}"]`);
        if (currentTh) {
            currentTh.classList.add(`sort-${currentSort.order}`);
            currentTh.setAttribute('data-order', currentSort.order);
        }
    }

    async function refreshDigest() {
        try {
            const res = await fetch(`${SITE_URL}/_api/contextinfo`, { 
                method: "POST", 
                headers: { "Accept": "application/json;odata=nometadata" },
                credentials: 'same-origin'
            });
            const data = await res.json();
            requestDigest = data.FormDigestValue;
            setTimeout(refreshDigest, 1500000);
        } catch (e) {
            setTimeout(refreshDigest, 300000);
        }
    }

    async function fetchType() {
        try {
            const r = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')?$select=ListItemEntityTypeFullName`, {
                headers: { "Accept": "application/json;odata=nometadata" },
                credentials: 'same-origin'
            });
            const d = await r.json();
            listType = d.ListItemEntityTypeFullName;
        } catch(e) {
            console.warn('Could not fetch list type, metadata will be omitted:', e);
        }
    }

    async function loadData() {
        try {
            const r = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items?$select=Id,EmployeeName,FirstName,LastName,M365Username_x002f_Computerusern,Windows_x002f_M365Password,Five9Username,Five9Password,Five9StationID,ILOANID,Department,Position,HireDate,Location,Manager,PhoneNumber,ExternalEmail,StatusOfNewHire,CreateFolder,CompanyBookings&$top=5000`, { 
                headers: { "Accept": "application/json;odata=nometadata" },
                credentials: 'same-origin'
            });
            
            if (!r.ok) {
                throw new Error(`HTTP error! status: ${r.status}`);
            }
            
            const data = await r.json();
            const allData = data.value || [];
            
            masterData = allData.filter(item => 
                item.StatusOfNewHire === 'Pending' || item.StatusOfNewHire === 'Completed'
            );
            
            allDepartments.clear();
            allLocations.clear();
            
            allData.forEach(item => {
                if (item.Department && item.Department.trim() !== '') {
                    allDepartments.add(item.Department);
                }
                if (item.Location && item.Location.trim() !== '') {
                    allLocations.add(item.Location);
                }
            });
            
            selectedIds.clear();
            
            populateFilterOptions();
            
            updateStatistics();
            updateDynamicFilters();
            
        } catch (error) {
            console.error('Error loading data:', error);
            showStatus('Error loading data from SharePoint', 'error');
        }
    }

    async function refreshDataPreserveState() {
        const savedSelections = Array.from(selectedIds);
        const savedFilters = { ...currentFilters };
        
        await loadData();
        
        currentFilters = savedFilters;
        const fNameEl = document.getElementById('fName');
        const fDeptEl = document.getElementById('fDept');
        const fLocationEl = document.getElementById('fLocation');
        const fHireDateEl = document.getElementById('fHireDate');
        const fStatusEl = document.getElementById('fStatus');
        
        if (fNameEl) fNameEl.value = savedFilters.fName || '';
        if (fDeptEl) fDeptEl.value = savedFilters.fDept || '';
        if (fLocationEl) fLocationEl.value = savedFilters.fLocation || '';
        if (fHireDateEl) fHireDateEl.value = savedFilters.fHireDate || '';
        if (fStatusEl) fStatusEl.value = savedFilters.fStatus || 'pending';
        
        selectedIds.clear();
        for (let id of savedSelections) {
            const stillExists = masterData.some(record => record.Id === id);
            if (stillExists) {
                selectedIds.add(id);
            }
        }
        
        updateDynamicFilters();
        
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
    }

    function updateStatistics() {
        const totalCountEl = document.getElementById('totalCount');
        const monthCountEl = document.getElementById('monthCount');
        const deptCountEl = document.getElementById('deptCount');
        const locationCountEl = document.getElementById('locationCount');
        const completedCountEl = document.getElementById('completedCount');
        const pendingCountEl = document.getElementById('pendingCount');
        const csCountEl = document.getElementById('csCount');
        const assessmentCountEl = document.getElementById('assessmentCount');
        const collectionsCountEl = document.getElementById('collectionsCount');
        const itCountEl = document.getElementById('itCount');
        const csDetailsEl = document.getElementById('csDetails');
        const assessmentDetailsEl = document.getElementById('assessmentDetails');
        const collectionsDetailsEl = document.getElementById('collectionsDetails');
        const itDetailsEl = document.getElementById('itDetails');
        
        if (!totalCountEl) {
            console.error('Required DOM elements not found');
            return;
        }
        
        const totalCount = masterData.length;
        
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();
        const monthCount = masterData.filter(item => {
            if (!item.HireDate) return false;
            const hireDate = new Date(item.HireDate);
            return hireDate.getMonth() === currentMonth && hireDate.getFullYear() === currentYear;
        }).length;
        
        const deptCount = allDepartments.size;
        const locationCount = allLocations.size;
        
        const completedCount = masterData.filter(item => item.StatusOfNewHire === 'Completed').length;
        const pendingCount = masterData.filter(item => item.StatusOfNewHire === 'Pending').length;
        
        const csCount = masterData.filter(item => 
            item.Department && item.Department.toLowerCase().includes('customer')).length;
        const assessmentCount = masterData.filter(item => 
            item.Department && item.Department.toLowerCase().includes('assessment')).length;
        const collectionsCount = masterData.filter(item => 
            item.Department && item.Department.toLowerCase().includes('collection')).length;
        const itCount = masterData.filter(item => 
            item.Department && (item.Department.toLowerCase() === 'it' || 
            item.Department.toLowerCase().includes('information technology'))).length;
        
        if (totalCountEl) totalCountEl.textContent = totalCount;
        if (monthCountEl) monthCountEl.textContent = monthCount;
        if (deptCountEl) deptCountEl.textContent = deptCount;
        if (locationCountEl) locationCountEl.textContent = locationCount;
        if (completedCountEl) completedCountEl.textContent = completedCount;
        if (pendingCountEl) pendingCountEl.textContent = pendingCount;
        
        if (csCountEl) csCountEl.textContent = csCount;
        if (assessmentCountEl) assessmentCountEl.textContent = assessmentCount;
        if (collectionsCountEl) collectionsCountEl.textContent = collectionsCount;
        if (itCountEl) itCountEl.textContent = itCount;
        
        if (csDetailsEl) csDetailsEl.innerHTML = getDepartmentDetails('Customer Service');
        if (assessmentDetailsEl) assessmentDetailsEl.innerHTML = getDepartmentDetails('Assessment');
        if (collectionsDetailsEl) collectionsDetailsEl.innerHTML = getDepartmentDetails('Collections');
        if (itDetailsEl) itDetailsEl.innerHTML = getDepartmentDetails('IT');
        
        renderLocationStats();
    }

    function getDepartmentDetails(department) {
        const deptItems = masterData.filter(item => {
            if (!item.Department) return false;
            const dept = item.Department.toLowerCase();
            if (department === 'Customer Service') {
                return dept.includes('customer');
            } else if (department === 'Assessment') {
                return dept.includes('assessment');
            } else if (department === 'Collections') {
                return dept.includes('collection');
            } else if (department === 'IT') {
                return dept === 'it' || dept.includes('information technology');
            }
            return false;
        });
        
        if (deptItems.length === 0) return '';
        
        const pending = deptItems.filter(item => item.StatusOfNewHire === 'Pending').length;
        const completed = deptItems.filter(item => item.StatusOfNewHire === 'Completed').length;
        
        const pendingDates = {};
        deptItems.filter(item => item.StatusOfNewHire === 'Pending').forEach(item => {
            if (item.HireDate) {
                const date = formatDateForDisplay(item.HireDate);
                pendingDates[date] = (pendingDates[date] || 0) + 1;
            }
        });
        
        let pendingHtml = '';
        Object.entries(pendingDates).forEach(([date, count]) => {
            pendingHtml += `<div class="pending-info">Pending ${date}: ${count}</div>`;
        });
        
        return `
            <div class="dept-breakdown">
                <div class="dept-item">
                    <span class="dept-name">Pending:</span>
                    <span class="pending-info">${pending}</span>
                </div>
                <div class="dept-item">
                    <span class="dept-name">Completed:</span>
                    <span class="completed-info">${completed}</span>
                </div>
                ${pendingHtml}
            </div>
        `;
    }

    function renderLocationStats() {
        const container = document.getElementById('locationStatsContainer');
        if (!container) {
            console.error('Location stats container not found');
            return;
        }
        
        const locations = {};
        
        masterData.forEach(item => {
            if (!item.Location) return;
            
            if (!locations[item.Location]) {
                locations[item.Location] = {
                    total: 0,
                    pending: 0,
                    completed: 0,
                    departments: {},
                    pendingDates: {}
                };
            }
            
            locations[item.Location].total++;
            
            if (item.StatusOfNewHire === 'Pending') {
                locations[item.Location].pending++;
                if (item.HireDate) {
                    const date = formatDateForDisplay(item.HireDate);
                    locations[item.Location].pendingDates[date] = 
                        (locations[item.Location].pendingDates[date] || 0) + 1;
                }
            } else if (item.StatusOfNewHire === 'Completed') {
                locations[item.Location].completed++;
            }
            
            const dept = item.Department || 'Unknown';
            if (!locations[item.Location].departments[dept]) {
                locations[item.Location].departments[dept] = {
                    total: 0,
                    pending: 0,
                    completed: 0
                };
            }
            
            locations[item.Location].departments[dept].total++;
            if (item.StatusOfNewHire === 'Pending') {
                locations[item.Location].departments[dept].pending++;
            } else if (item.StatusOfNewHire === 'Completed') {
                locations[item.Location].departments[dept].completed++;
            }
        });
        
        let html = '';
        
        Object.keys(locations).sort().forEach(location => {
            const loc = locations[location];
            
            let deptHtml = '';
            Object.keys(loc.departments).sort().forEach(dept => {
                const d = loc.departments[dept];
                deptHtml += `
                    <div class="dept-item">
                        <span class="dept-name">${escapeHtml(dept)}:</span>
                        <span>${d.total} (P:${d.pending}/C:${d.completed})</span>
                    </div>
                `;
            });
            
            let pendingHtml = '';
            Object.entries(loc.pendingDates).forEach(([date, count]) => {
                pendingHtml += `<div class="pending-info">Pending ${date}: ${count}</div>`;
            });
            
            html += `
                <div class="stat-card">
                    <div class="stat-title">${escapeHtml(location)}</div>
                    <div class="stat-value">${loc.total}</div>
                    <div class="stat-desc">Total employees</div>
                    <div class="location-details">
                        <div class="dept-breakdown">
                            <div class="dept-item">
                                <span class="dept-name">Pending:</span>
                                <span class="pending-info">${loc.pending}</span>
                            </div>
                            <div class="dept-item">
                                <span class="dept-name">Completed:</span>
                                <span class="completed-info">${loc.completed}</span>
                            </div>
                            ${deptHtml}
                            ${pendingHtml}
                        </div>
                    </div>
                </div>
            `;
        });
        
        container.innerHTML = html;
    }

    function normalizeHeader(h) {
        return h.toLowerCase()
            .replace(/\s+/g, '')
            .replace(/[^a-z0-9]/g, "");
    }

    function parseCsv(text) {
        const firstLine = text.split('\n')[0];
        const isTabDelimited = firstLine.includes('\t') && !firstLine.includes(',');
        
        const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
        if (!lines.length) return [];

        const delimiter = isTabDelimited ? '\t' : ',';
        
        function splitLine(line) {
            if (isTabDelimited) {
                return line.split(delimiter).map(v => v.trim().replace(/^"|"$/g, ''));
            } else {
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
                    } else if (ch === delimiter && !inQuotes) {
                        result.push(current);
                        current = "";
                    } else {
                        current += ch;
                    }
                }
                result.push(current);
                return result.map(v => v.trim().replace(/^"|"$/g, ''));
            }
        }

        const rawHeaders = splitLine(lines[0]);
        const headerNorm = rawHeaders.map(normalizeHeader);

        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            const cells = splitLine(lines[i]);
            if (cells.length === 1 && cells[0] === "") continue;

            const row = {};
            for (let c = 0; c < rawHeaders.length; c++) {
                const key = headerNorm[c];
                row[key] = cells[c] || "";
            }
            rows.push(row);
        }
        return rows;
    }

    function mapCsvRowToPayload(normRow) {
        const fullName = (normRow["fullname"] || normRow["full name"] || normRow["name"] || "").trim();
        const { firstName, lastName } = extractNames(fullName);
        
        const location = (normRow["location"] || "").trim();
        const position = (normRow["jobtitle"] || normRow["job title"] || normRow["position"] || "").trim();
        const department = (normRow["department"] || "").trim();
        const rawPhone = (normRow["phonenumber"] || normRow["phone number"] || "").trim();
        const externalEmail = (normRow["externalemail"] || normRow["external email"] || normRow["externalid"] || normRow["externalid"] || "").trim();
        const hireDate = (normRow["hiredate"] || normRow["hire date"] || "").trim();
        const manager = (normRow["manager"] || "").trim();
        
        const formattedPhone = formatPhoneNumber(rawPhone, location);
        
        return {
            EmployeeName: fullName,
            FirstName: firstName,
            LastName: lastName,
            Location: location,
            Department: department,
            Position: position,
            PhoneNumber: formattedPhone,
            M365Username_x002f_Computerusern: generateEmailUsername(firstName, lastName),
            Windows_x002f_M365Password: "YOUR_DEFAULT_PASSWORD",
            Five9Username: `${firstName.toLowerCase().replace(/[^a-z]/g, '')}.${lastName.toLowerCase().replace(/[^a-z]/g, '')}`,
            Five9Password: "YOUR_DEFAULT_PASSWORD",
            Five9StationID: "0",
            ILOANID: generateIloanId(firstName, lastName),
            HireDate: formatDateForSharePoint(hireDate),
            Manager: manager,
            ExternalEmail: externalEmail,
            StatusOfNewHire: "Pending",
            CreateFolder: null,
            CompanyBookings: null
        };
    }

    function getValidationErrors(payload) {
        const errors = [];
        if (!payload.EmployeeName || payload.EmployeeName.trim() === "") errors.push("Name is missing");
        if (!payload.Location || payload.Location.trim() === "") errors.push("Location is missing");
        if (!payload.Department || payload.Department.trim() === "") errors.push("Department is missing");
        if (!payload.Position || payload.Position.trim() === "") errors.push("Position is missing");
        if (!payload.HireDate) errors.push("Hire Date is missing or invalid");
        if (isEmailRequired(payload.Location)) {
            if (!payload.ExternalEmail || payload.ExternalEmail.trim() === "") {
                errors.push("External Email is required for this location");
            } else if (!payload.ExternalEmail.includes('@')) {
                errors.push("External Email format is invalid");
            }
        }
        return errors;
    }

    function isValidNewHireRow(payload) {
        return getValidationErrors(payload).length === 0;
    }

    function updateAutoGeneratedFields(index) {
        if (!csvRows[index]) return;
        
        const row = csvRows[index];
        const fullName = row.payload.EmployeeName || "";
        const location = row.payload.Location || "";
        
        const { firstName, lastName } = extractNames(fullName);
        
        row.payload.FirstName = firstName;
        row.payload.LastName = lastName;
        row.payload.M365Username_x002f_Computerusern = generateEmailUsername(firstName, lastName);
        row.payload.Five9Username = `${firstName.toLowerCase().replace(/[^a-z]/g, '')}.${lastName.toLowerCase().replace(/[^a-z]/g, '')}`;
        row.payload.ILOANID = generateIloanId(firstName, lastName);
        
        if (row.payload.PhoneNumber) {
            row.payload.PhoneNumber = formatPhoneNumber(row.payload.PhoneNumber, location);
        }
    }

    async function handleCSV(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = async (ev) => {
            document.getElementById('csvLog').classList.add('hidden');
            document.getElementById('csvLog').textContent = "";
            
            const text = ev.target.result;
            const parsed = parseCsv(text);
            
            if (parsed.length === 0) {
                document.getElementById('csvSummary').textContent = "No valid rows found in CSV.";
                return;
            }
            
            logCsv(`Reading CSV: ${file.name}`);
            logCsv(`Found ${parsed.length} rows`);
            
            csvRows = [];
            
            for (let idx = 0; idx < parsed.length; idx++) {
                const normRow = parsed[idx];
                const rawPhone = (normRow["phonenumber"] || "").trim();
                const payload = mapCsvRowToPayload(normRow);

                const validationErrors = getValidationErrors(payload);
                const isValid = validationErrors.length === 0;

                // Name format check — first + last only
                const nameParts = (payload.EmployeeName || "").trim().split(/\s+/).filter(p => p.length > 0);
                if (nameParts.length > 2) {
                    logCsv(`Row ${idx + 1}: ⚠ Name "${payload.EmployeeName}" has more than 2 parts — First + Last name only, no middle names. Generated fields (username, ILOAN ID) may be incorrect.`);
                }

                // Phone number validation
                if (rawPhone) {
                    const allDigits = rawPhone.replace(/\D/g, '');
                    let phoneCountryCode = "+1";
                    if (payload.Location) {
                        const locLower = payload.Location.toLowerCase();
                        for (const [key, code] of Object.entries(locationCountryCode)) {
                            if (locLower.includes(key)) { phoneCountryCode = code; break; }
                        }
                    }
                    // Strip country code prefix when phone is already in +CC format
                    let localDigits = allDigits;
                    if (rawPhone.trimStart().startsWith('+')) {
                        const ccNum = phoneCountryCode.slice(1);
                        if (allDigits.startsWith(ccNum)) {
                            localDigits = allDigits.slice(ccNum.length);
                        }
                    }
                    const phoneCheck = validatePhoneForCountry(localDigits, phoneCountryCode, rawPhone);
                    if (!phoneCheck.valid) {
                        logCsv(`Row ${idx + 1}: ⚠ Phone — ${phoneCheck.message}`);
                    }
                }

                let existingId = null;
                if (isValid) {
                    const existing = masterData.find(item =>
                        item.EmployeeName === payload.EmployeeName ||
                        (item.ExternalEmail && item.ExternalEmail === payload.ExternalEmail)
                    );
                    existingId = existing ? existing.Id : null;
                }

                csvRows.push({
                    index: idx,
                    payload,
                    existingId,
                    include: isValid && !existingId,
                    editing: false,
                    isValid: isValid
                });

                if (!isValid) {
                    logCsv(`Row ${idx + 1}: SKIPPED — "${payload.EmployeeName || 'Unknown'}" — ${validationErrors.join(", ")}`);
                } else if (existingId) {
                    logCsv(`Row ${idx + 1}: EXISTING (ID ${existingId}) - "${payload.EmployeeName}"`);
                } else {
                    logCsv(`Row ${idx + 1}: VALID - "${payload.EmployeeName}"`);
                }
            }
            
            renderCsvPreview();
            document.getElementById('btnStart').classList.remove('hidden');

            const duplicates = csvRows.filter(r => r.isValid && r.existingId);
            if (duplicates.length > 0) {
                showDuplicateAlert(duplicates);
            }
        };
        reader.readAsText(file);
    }

    function renderCsvPreview() {
        const body = document.getElementById('csvPreviewBody');
        const summary = document.getElementById('csvSummary');
        
        if (!csvRows.length) {
            body.innerHTML = "";
            summary.textContent = "No rows loaded.";
            return;
        }
        
        let newCount = 0, existingCount = 0, invalidCount = 0, includedCount = 0;
        
        const rowsHtml = csvRows.map((r, i) => {
            if (!r.isValid) {
                invalidCount++;
                r.include = false;
            } else {
                if (r.existingId) existingCount++; else newCount++;
                if (r.include) includedCount++;
            }
            
            let rowClass = "";
            if (!r.isValid) rowClass = "missing-critical";
            else if (r.existingId) rowClass = "existing-row";
            
            const editRowClass = r.editing ? "edit-row" : "";
            
            if (!r.editing) {
                return `
                    <tr class="${rowClass} ${editRowClass}">
                        <td>${i + 1}</td>
                        <td>
                            <input type="checkbox" class="csv-row-check" ${r.include ? "checked" : ""} ${!r.isValid ? "disabled" : ""} data-index="${i}" />
                        </td>
                        <td>
                            ${!r.isValid ? 
                                '<span style="color: #a4262c">Invalid</span>' : 
                                `<span style="color: ${r.existingId ? "#0078d4" : "#107c10"}">
                                    ${r.existingId ? 'Existing' : 'New'}
                                </span>`
                            }
                        </td>
                        <td>${escapeHtml(r.payload.EmployeeName || "")}</td>
                        <td>${escapeHtml(r.payload.Location || "")}</td>
                        <td>${escapeHtml(r.payload.Department || "")}</td>
                        <td>${escapeHtml(r.payload.Position || "")}</td>
                        <td>${escapeHtml(r.payload.PhoneNumber || "")}</td>
                        <td>${escapeHtml(r.payload.ExternalEmail || "")}</td>
                        <td>${escapeHtml(formatDateForDisplay(r.payload.HireDate) || "")}</td>
                        <td>${escapeHtml(r.payload.Manager || "")}</td>
                        <td>
                            <div style="font-size: 10px;">
                                <strong>First:</strong> ${escapeHtml(r.payload.FirstName || "")}<br>
                                <strong>Last:</strong> ${escapeHtml(r.payload.LastName || "")}<br>
                                <strong>M365:</strong> ${escapeHtml(r.payload.M365Username_x002f_Computerusern || "")}<br>
                                <strong>ILOAN:</strong> ${escapeHtml(r.payload.ILOANID || "")}
                            </div>
                        </td>
                        <td>
                            <button class="btn btn-outline btn-xs edit-csv-row-btn" data-index="${i}" ${!r.isValid ? "disabled" : ""}>Edit</button>
                        </td>
                    </tr>
                `;
            } else {
                return `
                    <tr class="${editRowClass}">
                        <td>${i + 1}</td>
                        <td>
                            <input type="checkbox" class="csv-row-check" ${r.include ? "checked" : ""} data-index="${i}" />
                        </td>
                        <td>
                            <span style="color: ${r.existingId ? "#0078d4" : "#107c10"}">
                                ${r.existingId ? 'Existing' : 'New'}
                            </span>
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="EmployeeName" 
                                       value="${escapeHtml(r.payload.EmployeeName || "")}" 
                                       style="${!r.payload.EmployeeName ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="Location" 
                                       value="${escapeHtml(r.payload.Location || "")}" 
                                       style="${!r.payload.Location ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="Department" 
                                       value="${escapeHtml(r.payload.Department || "")}" 
                                       style="${!r.payload.Department ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="Position" 
                                       value="${escapeHtml(r.payload.Position || "")}" 
                                       style="${!r.payload.Position ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="PhoneNumber" 
                                       value="${escapeHtml(r.payload.PhoneNumber || "")}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="ExternalEmail" 
                                       value="${escapeHtml(r.payload.ExternalEmail || "")}" 
                                       style="${isEmailRequired(r.payload.Location) && !r.payload.ExternalEmail ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input type="date" class="inline-input edit-field" data-index="${i}" data-field="HireDate" 
                                       value="${r.payload.HireDate ? new Date(r.payload.HireDate).toISOString().split('T')[0] : ''}" 
                                       style="${!r.payload.HireDate ? 'border-color: #ff4d4d;' : ''}" />
                        </td>
                        <td>
                            <input class="inline-input edit-field" data-index="${i}" data-field="Manager" 
                                       value="${escapeHtml(r.payload.Manager || "")}" />
                        </td>
                        <td>
                            <div style="font-size: 10px; padding: 4px; background: #f5f5f5;">
                                <strong>Auto:</strong> ${escapeHtml(r.payload.FirstName || "")} ${escapeHtml(r.payload.LastName || "")}<br>
                                <small>M365: ${escapeHtml(r.payload.M365Username_x002f_Computerusern || "")}</small>
                            </div>
                        </td>
                        <td>
                            <button class="btn btn-primary btn-xs save-csv-row-btn" data-index="${i}">Save</button>
                            <button class="btn btn-outline btn-xs cancel-csv-row-btn" data-index="${i}">Cancel</button>
                        </td>
                    </tr>
                `;
            }
        }).join('');
        
        body.innerHTML = rowsHtml;
        
        summary.innerHTML = `Loaded ${csvRows.length} row(s). Included: ${includedCount}, New: ${newCount}, Existing: ${existingCount}${invalidCount > 0 ? `, Invalid: ${invalidCount}` : ''}`;
        
        attachCsvPreviewEvents();
        updateCsvSelectAllCheckbox();
    }

    function attachCsvPreviewEvents() {
        document.querySelectorAll('.csv-row-check').forEach(checkbox => {
            const index = parseInt(checkbox.getAttribute('data-index'));
            checkbox.addEventListener('change', () => {
                if (csvRows[index] && csvRows[index].isValid) {
                    csvRows[index].include = checkbox.checked;
                    updateCsvSelectAllCheckbox();
                }
            });
        });
        
        document.querySelectorAll('.edit-csv-row-btn').forEach(btn => {
            const index = parseInt(btn.getAttribute('data-index'));
            btn.addEventListener('click', () => {
                if (csvRows[index] && csvRows[index].isValid) {
                    csvRows[index].editing = true;
                    renderCsvPreview();
                }
            });
        });
        
        document.querySelectorAll('.save-csv-row-btn').forEach(btn => {
            const index = parseInt(btn.getAttribute('data-index'));
            btn.addEventListener('click', () => {
                if (csvRows[index]) {
                    csvRows[index].editing = false;
                    updateAutoGeneratedFields(index);
                    
                    const isValid = isValidNewHireRow(csvRows[index].payload);
                    csvRows[index].isValid = isValid;
                    csvRows[index].include = isValid && !csvRows[index].existingId;
                    
                    if (csvRows[index].payload.HireDate) {
                        csvRows[index].payload.HireDate = formatDateForSharePoint(csvRows[index].payload.HireDate);
                    }
                    
                    renderCsvPreview();
                }
            });
        });
        
        document.querySelectorAll('.cancel-csv-row-btn').forEach(btn => {
            const index = parseInt(btn.getAttribute('data-index'));
            btn.addEventListener('click', () => {
                if (csvRows[index]) {
                    csvRows[index].editing = false;
                    renderCsvPreview();
                }
            });
        });
        
        document.querySelectorAll('.edit-field').forEach(input => {
            input.addEventListener('input', () => {
                const index = parseInt(input.getAttribute('data-index'));
                const field = input.getAttribute('data-field');
                if (csvRows[index]) {
                    if (field === 'HireDate' && input.type === 'date') {
                        csvRows[index].payload[field] = input.value ? new Date(input.value).toISOString() : null;
                    } else {
                        csvRows[index].payload[field] = input.value.trim();
                    }
                    
                    if (!input.value.trim() && field !== 'PhoneNumber' && field !== 'ExternalEmail') {
                        input.style.borderColor = '#ff4d4d';
                    } else {
                        input.style.borderColor = '#ddd';
                    }
                }
            });
        });
    }

    function logCsv(msg) {
        const logEl = document.getElementById('csvLog');
        logEl.classList.remove('hidden');
        logEl.textContent += msg + "\n";
        logEl.scrollTop = logEl.scrollHeight;
    }

    async function importCsvRows() {
        const selectedRows = csvRows.filter(r => r.include && r.isValid);
        
        if (selectedRows.length === 0) {
            showStatus("No valid rows selected for import.", "error");
            return;
        }
        
        if (!requestDigest) await refreshDigest();
        
        logCsv(`Starting import of ${selectedRows.length} selected rows...`);
        
        let success = 0, fail = 0;
        
        for (const row of selectedRows) {
            const name = row.payload.EmployeeName;
            const i = row.index;
            
            logCsv(`Row ${i + 1}: Importing "${name}"`);
            
            try {
                if (row.existingId) {
                    const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${row.existingId})`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=nometadata",
                            "IF-MATCH": "*",
                            "X-HTTP-Method": "MERGE",
                            "X-RequestDigest": requestDigest
                        },
                        credentials: 'same-origin',
                        body: JSON.stringify(row.payload)
                    });
                    
                    if (resp.ok) {
                        logCsv(`✓ Row ${i + 1}: Updated existing "${name}"`);
                        success++;
                    } else {
                        logCsv(`✗ Row ${i + 1}: ERROR updating "${name}" (HTTP ${resp.status})`);
                        fail++;
                    }
                } else {
                    const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=nometadata",
                            "X-RequestDigest": requestDigest
                        },
                        credentials: 'same-origin',
                        body: JSON.stringify(row.payload)
                    });
                    
                    if (resp.ok) {
                        const data = await resp.json();
                        logCsv(`✓ Row ${i + 1}: Created new "${name}" (ID ${data.Id})`);
                        success++;
                    } else {
                        logCsv(`✗ Row ${i + 1}: ERROR creating "${name}" (HTTP ${resp.status})`);
                        fail++;
                    }
                }
            } catch (e) {
                logCsv(`✗ Row ${i + 1}: SYSTEM ERROR for "${name}"`);
                fail++;
            }
        }
        
        logCsv(`Import complete. Success: ${success}, Failed: ${fail}.`);
        showStatus(`Import complete: ${success} succeeded, ${fail} failed`, fail === 0 ? "success" : "warning");
        await refreshDataPreserveState();
        
        document.getElementById('csvFile').value = '';
        csvRows = [];
        renderCsvPreview();
        document.getElementById('btnStart').classList.add('hidden');
    }

    function renderMain() {
        filteredData = applyFiltersToData(masterData, currentFilters);

        const sortFieldMap = {
            'employeename': 'EmployeeName',
            'firstname': 'FirstName',
            'lastname': 'LastName',
            'hiredate': 'HireDate',
            'statusofnewhire': 'StatusOfNewHire'
        };
        
        const sortField = sortFieldMap[currentSort.field] || 'HireDate';
        const sortedData = sortData(filteredData, sortField, currentSort.order);

        document.getElementById('mainTbody').innerHTML = sortedData.map((i, index) => {
            const status = i.StatusOfNewHire || 'Pending';
            const createFolder = i.CreateFolder || '';
            const CompanyBookings = i.CompanyBookings || '';
            
            return `
                <tr data-id="${i.Id}">
                    <td>
                        <input type="checkbox" class="row-check" ${selectedIds.has(i.Id) ? 'checked' : ''}>
                        <span class="serial-number">${index + 1}</span>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.EmployeeName||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.FirstName||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.LastName||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.M365Username_x002f_Computerusern||'')}" readonly>
                    </td>
                    <td>
                        <input type="text" class="inline-input read-only" value="${escapeHtml(i.Windows_x002f_M365Password||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Five9Username||'')}" readonly>
                    </td>
                    <td>
                        <input type="text" class="inline-input read-only" value="${escapeHtml(i.Five9Password||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Five9StationID||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.ILOANID||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Department||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Position||'')}" readonly>
                    </td>
                    <td>
                        <input type="text" class="inline-input read-only" value="${escapeHtml(formatDateForDisplay(i.HireDate) || '')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Location||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.Manager||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.PhoneNumber||'')}" readonly>
                    </td>
                    <td>
                        <input class="inline-input read-only" value="${escapeHtml(i.ExternalEmail||'')}" readonly>
                    </td>
                    <td>
                        <select class="inline-input s-createfolder" disabled>
                            <option value="" ${createFolder === '' ? 'selected' : ''}>No</option>
                            <option value="Yes" ${createFolder === 'Yes' ? 'selected' : ''}>Yes</option>
                        </select>
                    </td>
                    <td>
                        <select class="inline-input s-CompanyBookings" disabled>
                            <option value="" ${CompanyBookings === '' ? 'selected' : ''}>No</option>
                            <option value="Yes" ${CompanyBookings === 'Yes' ? 'selected' : ''}>Yes</option>
                        </select>
                    </td>
                    <td>
                        <select class="inline-input s-status" disabled>
                            <option value="Pending" ${status === 'Pending' ? 'selected' : ''}>Pending</option>
                            <option value="Completed" ${status === 'Completed' ? 'selected' : ''}>Completed</option>
                        </select>
                    </td>
                    <td>
                        <button class="btn btn-outline btn-small btn-edit">Edit</button>
                        <button class="btn btn-primary btn-small btn-save hidden">Save</button>
                        <button class="btn btn-danger btn-small btn-del">Delete</button>
                    </td>
                </tr>`;
        }).join('');
        
        updateSortUI();
        attachRowEvents();
        updateFilterSummary();
    }

    function updateFilterSummary() {
        const summaryEl = document.getElementById('filterSummary');
        if (!summaryEl) return;

        const pendingInView = filteredData.filter(r => r.StatusOfNewHire === 'Pending');
        const uniqueDepts = new Set(pendingInView.map(r => r.Department).filter(Boolean));
        const uniquePositions = new Set(pendingInView.map(r => r.Position).filter(Boolean));

        if (filteredData.length === 0) {
            summaryEl.classList.add('hidden');
            return;
        }

        summaryEl.classList.remove('hidden');

        const recordsEl = document.getElementById('fSummaryRecords');
        const deptsEl = document.getElementById('fSummaryDepts');
        const positionsEl = document.getElementById('fSummaryPositions');
        const deptsPluralEl = document.getElementById('fSummaryDeptsPlural');
        const positionsPluralEl = document.getElementById('fSummaryPositionsPlural');

        if (recordsEl) recordsEl.textContent = pendingInView.length;
        if (deptsEl) deptsEl.textContent = uniqueDepts.size;
        if (positionsEl) positionsEl.textContent = uniquePositions.size;
        if (deptsPluralEl) deptsPluralEl.textContent = uniqueDepts.size === 1 ? '' : 's';
        if (positionsPluralEl) positionsPluralEl.textContent = uniquePositions.size === 1 ? '' : 's';
    }

    function attachRowEvents() {
        document.querySelectorAll('.row-check').forEach(cb => {
            cb.onchange = (e) => {
                const id = parseInt(e.target.closest('tr').dataset.id);
                e.target.checked ? selectedIds.add(id) : selectedIds.delete(id);
                updateBulkDeleteButton();
                updateEmailButton();
                updateMarkCompletedButton();
                updateFolderButtons();
            };
        });
        
        document.querySelectorAll('.btn-edit').forEach(btn => {
            btn.onclick = (e) => {
                const tr = e.target.closest('tr');
                tr.querySelectorAll('.inline-input').forEach(i => { 
                    i.readOnly = false; 
                    i.disabled = false;
                    i.classList.remove('read-only'); 
                });
                tr.querySelector('.btn-edit').classList.add('hidden');
                tr.querySelector('.btn-save').classList.remove('hidden');
            }
        });
        
        document.querySelectorAll('.btn-save').forEach(btn => {
            btn.onclick = async (e) => {
                const tr = e.target.closest('tr');
                const id = tr.dataset.id;
                
                const data = { 
                    EmployeeName: tr.querySelector('td:nth-child(2) .inline-input').value.trim(),
                    FirstName: tr.querySelector('td:nth-child(3) .inline-input').value.trim(),
                    LastName: tr.querySelector('td:nth-child(4) .inline-input').value.trim(),
                    M365Username_x002f_Computerusern: tr.querySelector('td:nth-child(5) .inline-input').value.trim(),
                    Windows_x002f_M365Password: tr.querySelector('td:nth-child(6) .inline-input').value.trim(),
                    Five9Username: tr.querySelector('td:nth-child(7) .inline-input').value.trim(),
                    Five9Password: tr.querySelector('td:nth-child(8) .inline-input').value.trim(),
                    Five9StationID: tr.querySelector('td:nth-child(9) .inline-input').value.trim(),
                    ILOANID: tr.querySelector('td:nth-child(10) .inline-input').value.trim(),
                    Department: tr.querySelector('td:nth-child(11) .inline-input').value.trim(),
                    Position: tr.querySelector('td:nth-child(12) .inline-input').value.trim(),
                    HireDate: formatDateForSharePoint(tr.querySelector('td:nth-child(13) .inline-input').value),
                    Location: tr.querySelector('td:nth-child(14) .inline-input').value.trim(),
                    Manager: tr.querySelector('td:nth-child(15) .inline-input').value.trim(),
                    PhoneNumber: tr.querySelector('td:nth-child(16) .inline-input').value.trim(),
                    ExternalEmail: tr.querySelector('td:nth-child(17) .inline-input').value.trim(),
                    CreateFolder: tr.querySelector('td:nth-child(18) select').value,
                    CompanyBookings: tr.querySelector('td:nth-child(19) select').value,
                    StatusOfNewHire: tr.querySelector('td:nth-child(20) select').value
                };
                
                if (!requestDigest) await refreshDigest();
                
                try {
                    const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                        method: "POST",
                        body: JSON.stringify(data),
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=nometadata",
                            "X-RequestDigest": requestDigest,
                            "X-HTTP-Method": "MERGE",
                            "IF-MATCH": "*"
                        },
                        credentials: 'same-origin'
                    });

                    if (resp.ok) {
                        const record = masterData.find(item => item.Id == id);
                        if (record) Object.assign(record, data);
                        showStatus("Record Updated", "success");
                        updateDynamicFilters();
                    } else {
                        showStatus("Error updating record", "error");
                    }
                } catch (error) {
                    console.error('Save error:', error);
                    showStatus("Error updating record", "error");
                }
            }
        });
        
        document.querySelectorAll('.btn-del').forEach(btn => {
            btn.onclick = (e) => {
                const tr = e.target.closest('tr');
                const id = tr.dataset.id;
                showConfirmation(id);
            };
        });
    }

    function toggleAllCheckboxes(e) {
        const filteredIds = Array.from(document.querySelectorAll('#mainTbody tr'))
            .map(tr => parseInt(tr.dataset.id))
            .filter(id => !isNaN(id));
            
        if(e.target.checked) {
            filteredIds.forEach(id => selectedIds.add(id));
        } else {
            filteredIds.forEach(id => selectedIds.delete(id));
        }
        
        document.querySelectorAll('.row-check').forEach(cb => {
            const id = parseInt(cb.closest('tr').dataset.id);
            cb.checked = selectedIds.has(id);
        });
        
        updateBulkDeleteButton();
        updateEmailButton();
        updateMarkCompletedButton();
        updateFolderButtons();
    }

    function updateBulkDeleteButton() {
        const btn = document.getElementById('btnDeleteSelected');
        if (btn) {
            btn.classList.toggle('hidden', selectedIds.size === 0);
            btn.innerText = `🗑️ Delete Selected (${selectedIds.size})`;
        }
    }

    function updateEmailButton() {
        const btn = document.getElementById('btnEmailSelected');
        if (btn) {
            btn.classList.toggle('hidden', selectedIds.size === 0);
            btn.innerText = `✉️ Email Selected (${selectedIds.size})`;
        }
    }

    function updateMarkCompletedButton() {
        const btn = document.getElementById('btnMarkCompleted');
        if (btn) {
            btn.classList.toggle('hidden', selectedIds.size === 0);
            btn.innerText = `✓ Mark Completed (${selectedIds.size})`;
        }
    }

    function updateFolderButtons() {
        const btnFolders = document.getElementById('btnCreateFolders');
        const btnBookings = document.getElementById('btnCreateBookings');
        
        if (btnFolders) {
            btnFolders.classList.toggle('hidden', selectedIds.size === 0);
            btnFolders.innerText = `📁 Create Folders (${selectedIds.size})`;
        }
        
        if (btnBookings) {
            btnBookings.classList.toggle('hidden', selectedIds.size === 0);
            btnBookings.innerText = `📅 Create Bookings Profile (${selectedIds.size})`;
        }

        updatePushToITButton();
        updateLMProfileCSVButton();
    }

    async function deleteSelected() {
        if(!confirm(`Delete ${selectedIds.size} selected records?`)) return;
        if (!requestDigest) await refreshDigest();
        
        let deleted = 0, failed = 0;
        
        for(let id of selectedIds) {
            try {
                const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                    method: "POST",
                    headers: {
                        "X-RequestDigest": requestDigest,
                        "X-HTTP-Method": "DELETE",
                        "IF-MATCH": "*"
                    },
                    credentials: 'same-origin'
                });
                if (resp.ok) deleted++; else failed++;
            } catch (error) {
                failed++;
            }
        }
        
        selectedIds.clear();
        showStatus(`Deleted ${deleted} records${failed > 0 ? `, ${failed} failed` : ''}`, failed === 0 ? "success" : "warning");
        await refreshDataPreserveState();
    }

    function updatePushToITButton() {
        const btn = document.getElementById('btnPushToIT');
        if (btn) {
            btn.classList.toggle('hidden', selectedIds.size === 0);
            btn.innerText = `📤 Push to IT List (${selectedIds.size})`;
        }
    }

    function updateLMProfileCSVButton() {
        const btn = document.getElementById('btnLMProfileCSV');
        if (btn) {
            btn.classList.toggle('hidden', selectedIds.size === 0);
            btn.innerText = `📊 LM Profile CSV (${selectedIds.size})`;
        }
    }

    function downloadLMProfileCSV() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        if (selectedRecords.length === 0) {
            showStatus("Please select records to download LM Profile CSV", "warning");
            return;
        }

        const headers = ['First Name', 'Last Name', 'Email', 'System Name', 'iLoan User ID', 'Five9 Name', 'Hire Date'];
        let csvContent = headers.join(',') + '\n';

        selectedRecords.forEach(record => {
            const row = [
                escapeCSVValue(record.FirstName),
                escapeCSVValue(record.LastName),
                escapeCSVValue(record.M365Username_x002f_Computerusern),
                escapeCSVValue(''),
                escapeCSVValue(record.ILOANID),
                escapeCSVValue(record.Five9Username),
                escapeCSVValue(formatDateForDisplay(record.HireDate))
            ];
            csvContent += row.join(',') + '\n';
        });

        const today = new Date().toISOString().split('T')[0];
        const filename = `LM_Profile_${today}_${selectedRecords.length}_records.csv`;
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });

        if (navigator.msSaveBlob) {
            navigator.msSaveBlob(blob, filename);
        } else {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            link.style.display = 'none';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        showStatus(`LM Profile CSV downloaded (${selectedRecords.length} records)`, "success");
    }

    function mapToITPayload(record) {
        return {
            FullName: record.EmployeeName || "",
            FirstName: record.FirstName || "",
            LastName: record.LastName || "",
            M365Username_x002f_Computerusern: record.M365Username_x002f_Computerusern || "",
            Windows_x002f_M365Password: record.Windows_x002f_M365Password || "",
            Five9Username: record.Five9Username || "",
            Five9Password: record.Five9Password || "",
            Five9StationID: record.Five9StationID || "",
            ILOANID: record.ILOANID || "",
            Department: record.Department || "",
            Position: record.Position || "",
            HireDate: record.HireDate || null,
            Location: record.Location || "",
            Manager: record.Manager || "",
            PhoneNumber: record.PhoneNumber || "",
            ExternalSystemId: record.ExternalEmail || ""
        };
    }

    async function pushToITList() {
        const selectedRecords = masterData.filter(item => selectedIds.has(item.Id));
        if (selectedRecords.length === 0) {
            showStatus("Please select records to push to IT list", "warning");
            return;
        }

        showStatus(`Connecting to your SharePoint site...`, "info");

        // Get a fresh digest for the IT site (different site = different digest)
        let itDigest = "";
        try {
            const digestResp = await fetch(`${IT_SITE_URL}/_api/contextinfo`, {
                method: "POST",
                headers: { "Accept": "application/json;odata=nometadata" },
                credentials: 'same-origin'
            });
            if (!digestResp.ok) throw new Error(`HTTP ${digestResp.status}`);
            const digestData = await digestResp.json();
            itDigest = digestData.FormDigestValue;
        } catch (e) {
            showStatus("Could not connect to your SharePoint site. Check that you have access to that site.", "error");
            return;
        }

        showStatus(`Pushing ${selectedRecords.length} record(s) to IT list...`, "info");

        let created = 0, updated = 0, failed = 0;

        for (const record of selectedRecords) {
            const payload = mapToITPayload(record);
            const name = record.EmployeeName || "Unknown";

            try {
                // Check if record already exists by M365 username
                let existingId = null;
                const m365 = payload.M365Username_x002f_Computerusern;
                if (m365) {
                    const safe = m365.replace(/'/g, "''");
                    const checkResp = await fetch(
                        `${IT_SITE_URL}/_api/web/lists(guid'${IT_LIST_GUID}')/items?$filter=M365Username_x002f_Computerusern eq '${safe}'&$top=1&$select=Id`,
                        { headers: { "Accept": "application/json;odata=nometadata" }, credentials: 'same-origin' }
                    );
                    if (checkResp.ok) {
                        const checkData = await checkResp.json();
                        if (checkData.value && checkData.value.length > 0) {
                            existingId = checkData.value[0].Id;
                        }
                    }
                }

                if (existingId) {
                    const resp = await fetch(`${IT_SITE_URL}/_api/web/lists(guid'${IT_LIST_GUID}')/items(${existingId})`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=nometadata",
                            "X-HTTP-Method": "MERGE",
                            "IF-MATCH": "*",
                            "X-RequestDigest": itDigest
                        },
                        credentials: 'same-origin',
                        body: JSON.stringify(payload)
                    });
                    if (resp.ok) updated++; else { failed++; console.error(`IT push update failed for "${name}": HTTP ${resp.status}`); }
                } else {
                    const resp = await fetch(`${IT_SITE_URL}/_api/web/lists(guid'${IT_LIST_GUID}')/items`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "Content-Type": "application/json;odata=nometadata",
                            "X-RequestDigest": itDigest
                        },
                        credentials: 'same-origin',
                        body: JSON.stringify(payload)
                    });
                    if (resp.ok) created++; else { failed++; console.error(`IT push create failed for "${name}": HTTP ${resp.status}`); }
                }
            } catch (e) {
                failed++;
                console.error(`IT push error for "${name}":`, e);
            }
        }

        const msg = `IT List push complete: ${created} created, ${updated} updated${failed > 0 ? `, ${failed} failed` : ''}`;
        showStatus(msg, failed === 0 ? "success" : "warning");
    }

    function switchTab(t) {
        document.getElementById('viewSection').classList.toggle('hidden', t !== 'view');
        document.getElementById('filterContainer').classList.toggle('hidden', t !== 'view');
        document.getElementById('bulkSection').classList.toggle('hidden', t !== 'bulk');
        
        document.getElementById('tabView').classList.toggle('active', t === 'view');
        document.getElementById('tabBulk').classList.toggle('active', t === 'bulk');
        
        document.getElementById('statsContainer').classList.toggle('hidden', t !== 'view');
    }

    function showStatus(m, c) {
        const s = document.getElementById('statusBox');
        s.innerText = m; 
        s.className = 'status-box ' + c; 
        s.style.display = 'block';
        setTimeout(() => s.style.display = 'none', 5000);
    }

    function escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str).replace(/[&<>"']/g, (c) => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        }[c] || c));
    }

    // Defer initialization until DOM and key elements are present (modern pages)
    whenDomReady(async () => {
        const ready = await waitForElements([
            '#viewSection',
            '#filterContainer',
            '#tabView',
            '#tabBulk'
        ], 8000);
        if (!ready) {
            console.warn('NewHireDashboard: required elements not found within timeout; attempting init anyway.');
        }
        init();
    });
})();