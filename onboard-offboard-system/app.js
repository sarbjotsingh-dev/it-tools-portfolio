(function() {
            // SharePoint configuration for Onboarding/Offboarding list
            const LIST_GUID = "YOUR-RECORDS-LIST-GUID";
            const SITE_URL = "https://your-tenant.sharepoint.com";
            
            // Internal column names mapping
            const COLUMN_MAPPING = {
                Name: "Name",                    // Name field (Title is the default)
                Department: "Department",
                Location: "Location",
                OnboardingDate: "OnboardingDate",
                OffboardingDate: "OffboardingDate",
                OffboardingReason: "OffboardingReason",
                Position: "Position",
                RehireFlag: "RehireFlag",
                WorkEmail: "WorkEmail"
            };
            
            let listType = ""; 
            let masterData = [];
            let csvRows = [];
            let selectedIds = new Set();
            let requestDigest = "";
            let currentSort = { field: 'Name', order: 'asc' };
            let filteredData = [];
            let clearRecordId = null;
            let selectedStatus = "";

            async function init() {
                bindEvents();
                await fetchType();
                await loadData();
                await refreshDigest();
                setupSorting();
            }

            function bindEvents() {
                document.getElementById('tabView').onclick = () => switchTab('view');
                document.getElementById('tabBulk').onclick = () => switchTab('bulk');
                document.getElementById('btnRefresh').onclick = loadData;
                document.getElementById('btnFindDuplicates').onclick = findDuplicateNames;
                document.getElementById('templateBtn').onclick = downloadTemplate;
                document.getElementById('csvFile').onchange = handleCSV;
                document.getElementById('btnStart').onclick = importCsvRows;
                document.getElementById('btnDeleteSelected').onclick = deleteSelected;
                document.getElementById('btnExportSelected').onclick = exportSelected;
                document.getElementById('selectAll').onclick = toggleAllCheckboxes;
                
                // Status filter events
                document.getElementById('statusFilter').onchange = applyStatusFilter;
                document.getElementById('clearStatusFilter').onclick = clearStatusFilter;
                
                // Confirmation dialog events
                document.getElementById('confirmationCancel').onclick = hideConfirmation;
                document.getElementById('confirmationConfirm').onclick = performClearRecord;
                document.getElementById('confirmationOverlay').onclick = hideConfirmation;
                
                // Filter inputs
                document.getElementById('fName').oninput = renderMain;
                ['fDepartment', 'fLocation', 'fPosition', 'fMonth'].forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.onchange = renderMain;
                });
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
                    let aVal, bVal;
                    
                    switch(field) {
                        case 'title':
                            aVal = a.Name || '';
                            bVal = b.Name || '';
                            break;
                        case 'department':
                            aVal = a.Department || '';
                            bVal = b.Department || '';
                            break;
                        case 'location':
                            aVal = a.Location || '';
                            bVal = b.Location || '';
                            break;
                        case 'onboardingdate':
                            aVal = a.OnboardingDate ? new Date(a.OnboardingDate) : new Date(0);
                            bVal = b.OnboardingDate ? new Date(b.OnboardingDate) : new Date(0);
                            return order === 'asc' ? aVal - bVal : bVal - aVal;
                        case 'offboardingdate':
                            aVal = a.OffboardingDate ? new Date(a.OffboardingDate) : new Date(0);
                            bVal = b.OffboardingDate ? new Date(b.OffboardingDate) : new Date(0);
                            return order === 'asc' ? aVal - bVal : bVal - aVal;
                        case 'position':
                            aVal = a.Position || '';
                            bVal = b.Position || '';
                            break;
                        default:
                            aVal = a.Name || '';
                            bVal = b.Name || '';
                    }
                    
                    aVal = String(aVal).toLowerCase();
                    bVal = String(bVal).toLowerCase();
                    
                    if (order === 'asc') {
                        return aVal.localeCompare(bVal);
                    } else {
                        return bVal.localeCompare(aVal);
                    }
                });
            }

            function updateSortUI() {
                document.querySelectorAll('.sortable').forEach(th => {
                    th.classList.remove('sort-asc', 'sort-desc');
                });
                
                const currentTh = document.querySelector(`.sortable[data-sort="${currentSort.field}"]`);
                if (currentTh) {
                    currentTh.classList.add(`sort-${currentSort.order}`);
                    currentTh.setAttribute('data-order', currentSort.order);
                }
            }

            // --- DIGEST MANAGEMENT ---
            async function refreshDigest() {
                try {
                    const res = await fetch(`${SITE_URL}/_api/contextinfo`, { 
                        method: "POST", 
                        headers: { "Accept": "application/json;odata=nometadata" } 
                    });
                    const data = await res.json();
                    requestDigest = data.FormDigestValue;
                    setTimeout(refreshDigest, 1500000);
                } catch (e) {
                    console.error("Error refreshing digest:", e);
                    setTimeout(refreshDigest, 300000);
                }
            }

            // --- SHAREPOINT CORE ---
            async function fetchType() {
                try {
                    const r = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')?$select=ListItemEntityTypeFullName`, { 
                        headers: { "Accept": "application/json;odata=nometadata" } 
                    });
                    const d = await r.json();
                    listType = d.ListItemEntityTypeFullName;
                } catch (e) {
                    console.error("Error fetching list type:", e);
                }
            }

            async function loadData() {
                try {
                    const r = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items?$select=Id,Name,Department,Location,OnboardingDate,OffboardingDate,OffboardingReason,Position,RehireFlag,WorkEmail&$top=5000`, { 
                        headers: { "Accept": "application/json;odata=nometadata" } 
                    });
                    const data = await r.json();
                    masterData = data.value || [];
                    populateFilterOptions();
                    renderMain();
                } catch (e) {
                    console.error("Error loading data:", e);
                    showStatus("Error loading data from SharePoint", "error");
                }
            }

            // --- HELPER FUNCTIONS ---
            function getMonthKey(dateString) {
                if (!dateString) return null;
                const date = new Date(dateString);
                if (isNaN(date.getTime())) return null;
                // Use UTC methods to avoid timezone offset issues
                return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
            }

            function getMonthLabel(monthKey) {
                const [year, month] = (monthKey || '').split('-');
                if (!year || !month) return '';
                // Create UTC date to avoid timezone shift issues
                const d = new Date(Date.UTC(parseInt(year), parseInt(month) - 1, 1));
                return d.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' });
            }

            function populateFilterOptions() {
                const departments = Array.from(new Set(masterData.map(i => (i.Department || '').trim()).filter(Boolean))).sort();
                const locations = Array.from(new Set(masterData.map(i => (i.Location || '').trim()).filter(Boolean))).sort();
                const positions = Array.from(new Set(masterData.map(i => (i.Position || '').trim()).filter(Boolean))).sort();
                const months = Array.from(new Set(masterData.flatMap(i => [getMonthKey(i.OnboardingDate), getMonthKey(i.OffboardingDate)].filter(Boolean)))).sort();

                const deptEl = document.getElementById('fDepartment');
                const locEl = document.getElementById('fLocation');
                const posEl = document.getElementById('fPosition');
                const monthEl = document.getElementById('fMonth');

                function fillSelect(el, options, defaultText) {
                    if (!el) return;
                    const current = el.value || '';
                    el.innerHTML = `<option value="">${defaultText}</option>` + options.map(v => `\n<option value="${v}">${v}</option>`).join('');
                    if (current && options.includes(current)) {
                        el.value = current;
                    }
                }

                fillSelect(deptEl, departments, 'All Departments');
                fillSelect(locEl, locations, 'All Locations');
                fillSelect(posEl, positions, 'All Positions');
                if (monthEl) {
                    const monthOptions = months.map(key => `${key}:${getMonthLabel(key)}`);
                    monthEl.innerHTML = '<option value="">All Months</option>' + monthOptions.map(m => {
                        const parts = m.split(':');
                        return `<option value="${parts[0]}">${parts[1]}</option>`;
                    }).join('');
                }
            }

            function updateStatistics(dataSource) {
                const data = Array.isArray(dataSource) ? dataSource : (filteredData.length ? filteredData : masterData);
                const totalCount = data.length;
                const today = new Date();

                const onboardedCount = data.filter(item => !item.OffboardingDate).length;
                const offboardedCount = data.filter(item => item.OffboardingDate).length;
                const rehireCount = data.filter(item => item.RehireFlag && item.RehireFlag.toLowerCase() === 'yes').length;

                const selectedMonth = document.getElementById('fMonth') ? document.getElementById('fMonth').value : '';
                const onboardingMonthCount = data.filter(item => selectedMonth && getMonthKey(item.OnboardingDate) === selectedMonth).length;
                const offboardingMonthCount = data.filter(item => selectedMonth && getMonthKey(item.OffboardingDate) === selectedMonth).length;

                const uniqueDepartmentsCount = new Set(data.map(i => (i.Department || '').trim()).filter(Boolean)).size;
                const uniqueLocationsCount = new Set(data.map(i => (i.Location || '').trim()).filter(Boolean)).size;

                document.getElementById('totalCount').textContent = totalCount;
                document.getElementById('onboardingCount').textContent = onboardedCount;
                document.getElementById('offboardingCount').textContent = offboardedCount;
                document.getElementById('rehireCount').textContent = rehireCount;
                document.getElementById('onboardingMonthCount').textContent = onboardingMonthCount;
                document.getElementById('offboardingMonthCount').textContent = offboardingMonthCount;
                document.getElementById('uniqueDepartments').textContent = uniqueDepartmentsCount;
                document.getElementById('uniqueLocations').textContent = uniqueLocationsCount;
            }

            // --- STATUS FILTER FUNCTIONS ---
            function applyStatusFilter() {
                const select = document.getElementById('statusFilter');
                selectedStatus = select.value;
                renderMain();
            }

            function clearStatusFilter() {
                const select = document.getElementById('statusFilter');
                select.value = '';
                selectedStatus = '';
                renderMain();
            }

            // --- FIND DUPLICATE NAMES ---
            function findDuplicateNames() {
                const nameMap = {};
                const duplicates = new Set();
                
                masterData.forEach(item => {
                    const name = item.Name;
                    if (name) {
                        if (!nameMap[name]) {
                            nameMap[name] = [];
                        }
                        nameMap[name].push(item);
                    }
                });
                
                const duplicateItemIds = [];
                
                Object.keys(nameMap).forEach(name => {
                    const items = nameMap[name];
                    if (items.length > 1) {
                        items.sort((a, b) => a.Id - b.Id);
                        for (let i = 1; i < items.length; i++) {
                            duplicateItemIds.push(items[i].Id);
                        }
                    }
                });
                
                if (duplicateItemIds.length === 0) {
                    showStatus("No duplicate names found", "success");
                    return;
                }
                
                selectedIds.clear();
                duplicateItemIds.forEach(id => selectedIds.add(id));
                
                document.querySelectorAll('.row-check').forEach(cb => {
                    const id = parseInt(cb.closest('tr').dataset.id);
                    cb.checked = selectedIds.has(id);
                });
                
                document.getElementById('selectAll').checked = false;
                updateBulkActionButtons();
                
                showStatus(`Found ${duplicateItemIds.length} duplicate name records (copies selected)`, "warning");
                
                // Highlight duplicates
                duplicateItemIds.forEach(id => {
                    const row = document.querySelector(`tr[data-id="${id}"]`);
                    if (row) {
                        row.classList.add('duplicate-highlight');
                        setTimeout(() => {
                            row.classList.remove('duplicate-highlight');
                        }, 5000);
                    }
                });
            }

            // --- TEMPLATE DOWNLOAD ---
            function downloadTemplate() {
                const today = new Date().toISOString().split('T')[0];
                const nextMonth = new Date();
                nextMonth.setMonth(nextMonth.getMonth() + 1);
                const nextMonthDate = nextMonth.toISOString().split('T')[0];
                
                const csvContent = `Name,Department,Location,OnboardingDate,OffboardingDate,OffboardingReason,Position,RehireFlag,WorkEmail
"John Smith","IT","New York","${today}","","","Developer","Yes","john.smith@company.com"
"Jane Doe","HR","Chicago","${today}","${nextMonthDate}","Relocation","HR Manager","No","jane.doe@company.com"
"Bob Johnson","Finance","Remote","","${today}","Retirement","Accountant","Pending","bob.johnson@company.com"
"Alice Brown","Marketing","Los Angeles","${today}","","","Marketing Specialist","Yes","alice.brown@company.com"`;
                
                const filename = `Onboarding_Offboarding_Template_${today}.csv`;
                
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
                
                showStatus("Template downloaded successfully", "success");
            }

            // --- EXPORT SELECTED ---
            function exportSelected() {
                if (selectedIds.size === 0) {
                    showStatus("No items selected for export", "error");
                    return;
                }
                
                const selectedItems = masterData.filter(item => selectedIds.has(item.Id));
                
                if (selectedItems.length === 0) {
                    showStatus("Selected items not found in data", "error");
                    return;
                }
                
                selectedItems.sort((a, b) => (a.Name || '').localeCompare(b.Name || ''));
                
                const headers = ['Name', 'Department', 'Location', 'OnboardingDate', 'OffboardingDate', 
                               'OffboardingReason', 'Position', 'RehireFlag', 'WorkEmail'];
                const csvRows = [];
                
                csvRows.push(headers.join(','));
                
                selectedItems.forEach(item => {
                    const row = [
                        `"${(item.Name || '').replace(/"/g, '""')}"`,
                        `"${(item.Department || '').replace(/"/g, '""')}"`,
                        `"${(item.Location || '').replace(/"/g, '""')}"`,
                        `"${formatDateForDisplay(item.OnboardingDate) || ''}"`,
                        `"${formatDateForDisplay(item.OffboardingDate) || ''}"`,
                        `"${(item.OffboardingReason || '').replace(/"/g, '""')}"`,
                        `"${(item.Position || '').replace(/"/g, '""')}"`,
                        `"${(item.RehireFlag || '').replace(/"/g, '""')}"`,
                        `"${(item.WorkEmail || '').replace(/"/g, '""')}"`
                    ];
                    csvRows.push(row.join(','));
                });
                
                const csvContent = csvRows.join('\n');
                const filename = `Onboarding_Offboarding_Export_${new Date().toISOString().split('T')[0]}.csv`;
                
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
                
                showStatus(`Exported ${selectedItems.length} selected items to CSV`, "success");
            }

            // --- VALIDATION FUNCTIONS ---
            function isValidName(name) {
                return name && name.trim() !== "";
            }

            function formatDateForDisplay(dateString) {
                if (!dateString) return "";
                try {
                    const date = new Date(dateString);
                    if (isNaN(date.getTime())) return dateString;
                    return date.toISOString().split('T')[0];
                } catch (e) {
                    return dateString;
                }
            }

            function formatDateForSharePoint(dateString) {
                if (!dateString) return null;
                try {
                    const date = new Date(dateString);
                    if (isNaN(date.getTime())) return null;
                    return date.toISOString();
                } catch (e) {
                    return null;
                }
            }

            function validateRehireFlag(value) {
                const val = (value || '').toLowerCase();
                if (val === 'yes' || val === 'no' || val === 'pending') return true;
                return false;
            }

            // --- CSV PARSING ---
            function splitCsvLine(line) {
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
                    } else if (ch === "," && !inQuotes) {
                        result.push(current);
                        current = "";
                    } else {
                        current += ch;
                    }
                }
                result.push(current);
                return result.map(v => v.trim().replace(/^"|"$/g, ''));
            }

            function normalizeHeader(h) {
                return h.toLowerCase().replace(/[^a-z0-9]/g, "");
            }

            function parseCsv(text) {
                const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
                if (!lines.length) return [];

                const rawHeaders = splitCsvLine(lines[0]);
                const headerNorm = rawHeaders.map(normalizeHeader);

                const rows = [];
                for (let i = 1; i < lines.length; i++) {
                    const cells = splitCsvLine(lines[i]);
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
                // Map CSV columns to SharePoint internal column names
                return {
                    Name: (normRow["name"] || "").trim(),
                    Department: (normRow["department"] || "").trim(),
                    Location: (normRow["location"] || "").trim(),
                    OnboardingDate: formatDateForSharePoint(normRow["onboardingdate"] || normRow["onboarding date"] || ""),
                    OffboardingDate: formatDateForSharePoint(normRow["offboardingdate"] || normRow["offboarding date"] || ""),
                    OffboardingReason: (normRow["offboardingreason"] || normRow["offboarding reason"] || "").trim(),
                    Position: (normRow["position"] || "").trim(),
                    RehireFlag: (normRow["rehireflag"] || normRow["rehire flag"] || "").trim(),
                    WorkEmail: (normRow["workemail"] || normRow["work email"] || "").trim()
                };
            }

            async function findExistingByName(name) {
                if (!name) return null;
                
                const safeName = name.replace(/'/g, "''");
                const endpoint = `${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items?$top=1&$filter=Name eq '${safeName}'`;
                
                try {
                    const resp = await fetch(endpoint, {
                        headers: { "Accept": "application/json;odata=nometadata" }
                    });
                    if (!resp.ok) return null;
                    
                    const data = await resp.json();
                    return data.value && data.value.length > 0 ? data.value[0].Id : null;
                } catch (e) {
                    console.error("Error checking existing Name:", e);
                    return null;
                }
            }

            // --- CSV HANDLING ---
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
                    logCsv(`Parsed rows (excluding header): ${parsed.length}`);
                    
                    csvRows = [];
                    
                    for (let idx = 0; idx < parsed.length; idx++) {
                        const normRow = parsed[idx];
                        const payload = mapCsvRowToPayload(normRow);
                        
                        const isValid = isValidName(payload.Title);
                        const existingId = isValid ? await findExistingByName(payload.Title) : null;
                        
                        const shouldInclude = isValid && !existingId;
                        
                        csvRows.push({
                            index: idx,
                            payload,
                            existingId,
                            include: shouldInclude,
                            editing: false,
                            isValid: isValid
                        });
                        
                        if (!isValid) {
                            logCsv(`Row ${idx + 1}: SKIPPED - Missing Name`);
                        } else {
                            const status = existingId ? `EXISTING (ID ${existingId}) - Unchecked by default` : "NEW";
                            logCsv(`Row ${idx + 1}: ${payload.Title} - ${status}`);
                        }
                    }
                    
                    renderCsvPreview();
                    document.getElementById('btnStart').classList.remove('hidden');
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
                
                let newCount = 0;
                let existingCount = 0;
                let missingCritical = 0;
                let includedCount = 0;
                
                const rowsHtml = csvRows.map((r, i) => {
                    const isCriticalMissing = !isValidName(r.payload.Title);
                    
                    if (isCriticalMissing) {
                        missingCritical++;
                        r.include = false;
                    } else {
                        if (r.existingId) existingCount++; else newCount++;
                        if (r.include) includedCount++;
                    }
                    
                    let rowClass = "";
                    if (isCriticalMissing) {
                        rowClass = "missing-critical";
                    } else if (r.existingId) {
                        rowClass = "existing-row";
                    }
                    
                    const editRowClass = r.editing ? "edit-row" : "";
                    
                    if (!r.editing) {
                        return `
                            <tr class="${rowClass} ${editRowClass}">
<td>${i + 1}</td>
<td>
	<input type="checkbox" ${r.include ? "checked" : ""} ${isCriticalMissing ? "disabled" : ""} data-index="${i}" />
</td>
<td>
                                    ${isCriticalMissing ? 
                                        '<span style="font-weight: 600; color: #a4262c">Missing Name</span>' : 
                                        `<span style="font-weight: 600; color: ${r.existingId ? "#0078d4" : "#107c10"}">
                                            ${r.existingId ? `Existing (ID ${r.existingId})` : "New"}
                                        </span>`
                                    }
                                </td>
<td>${escapeHtml(r.payload.Title || "")}</td>
<td>${escapeHtml(r.payload.Department || "")}</td>
<td>${escapeHtml(r.payload.Location || "")}</td>
<td>${escapeHtml(formatDateForDisplay(r.payload.OnboardingDate) || "")}</td>
<td>${escapeHtml(formatDateForDisplay(r.payload.OffboardingDate) || "")}</td>
<td>${escapeHtml(r.payload.OffboardingReason || "")}</td>
<td>${escapeHtml(r.payload.Position || "")}</td>
<td>${escapeHtml(r.payload.RehireFlag || "")}</td>
<td>${escapeHtml(r.payload.WorkEmail || "")}</td>
<td>
	<button style="font-size: 11px; padding: 4px 8px;" class="btn btn-outline edit-csv-row-btn" data-index="${i}" ${isCriticalMissing ? "disabled" : ""}>Edit</button>
</td>
</tr>
                        `;
                    } else {
                        // Edit mode
                        return `
                            <tr class="${editRowClass}">
<td>${i + 1}</td>
<td>
	<input type="checkbox" ${r.include ? "checked" : ""} data-index="${i}" />
</td>
<td>
	<span style="font-weight: 600; color: ${r.existingId ? "#0078d4" : "#107c10"}">
                                        ${r.existingId ? `Existing (ID ${r.existingId})` : "New"}
                                    </span>
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="Name" 
                                           value="${escapeHtml(r.payload.Name || "")}" 
                                           placeholder="Required" />
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="Department" 
                                           value="${escapeHtml(r.payload.Department || "")}" />
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="Location" 
                                           value="${escapeHtml(r.payload.Location || "")}" />
</td>
<td>
	<input type="date" class="inline-input edit-field" data-index="${i}" data-field="OnboardingDate" 
                                           value="${escapeHtml(formatDateForDisplay(r.payload.OnboardingDate) || "")}" />
</td>
<td>
	<input type="date" class="inline-input edit-field" data-index="${i}" data-field="OffboardingDate" 
                                           value="${escapeHtml(formatDateForDisplay(r.payload.OffboardingDate) || "")}" />
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="OffboardingReason" 
                                           value="${escapeHtml(r.payload.OffboardingReason || "")}" />
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="Position" 
                                           value="${escapeHtml(r.payload.Position || "")}" />
</td>
<td>
	<select class="inline-input edit-field" data-index="${i}" data-field="RehireFlag">
		<option value="">Select</option>
		<option value="Yes" ${r.payload.RehireFlag === 'Yes' ? 'selected' : ''}>Yes</option>
		<option value="No" ${r.payload.RehireFlag === 'No' ? 'selected' : ''}>No</option>
		<option value="Pending" ${r.payload.RehireFlag === 'Pending' ? 'selected' : ''}>Pending</option>
	</select>
</td>
<td>
	<input class="inline-input edit-field" data-index="${i}" data-field="WorkEmail" 
                                           value="${escapeHtml(r.payload.WorkEmail || "")}" />
</td>
<td>
	<button style="font-size: 10px; padding: 3px 6px; margin-bottom: 3px;" class="btn btn-primary save-csv-row-btn" data-index="${i}">Save</button>
	<button style="font-size: 10px; padding: 3px 6px;" class="btn btn-outline cancel-csv-row-btn" data-index="${i}">Cancel</button>
</td>
</tr>
                        `;
                    }
                }).join('');
                
                body.innerHTML = rowsHtml;
                
                let summaryText = `Loaded ${csvRows.length} row(s). Included: ${includedCount}, New: ${newCount}, Existing: ${existingCount}`;
                
                if (existingCount > 0) {
                    summaryText += ` <span style="color: #0078d4;">(Existing items unchecked by default)</span>`;
                }
                
                if (missingCritical > 0) {
                    summaryText += ` <span style="color: #a4262c;">(${missingCritical} missing Name - auto-excluded)</span>`;
                }
                
                summary.innerHTML = summaryText;
                
                attachCsvPreviewEvents();
            }

            function attachCsvPreviewEvents() {
                // Checkboxes
                document.querySelectorAll('input[type="checkbox"][data-index]').forEach(checkbox => {
                    const index = parseInt(checkbox.getAttribute('data-index'));
                    checkbox.addEventListener('change', () => {
                        if (csvRows[index] && csvRows[index].isValid) {
                            csvRows[index].include = checkbox.checked;
                        } else {
                            checkbox.checked = false;
                        }
                    });
                });
                
                // Edit buttons
                document.querySelectorAll('.edit-csv-row-btn').forEach(btn => {
                    const index = parseInt(btn.getAttribute('data-index'));
                    btn.addEventListener('click', () => {
                        if (csvRows[index] && csvRows[index].isValid) {
                            csvRows[index].editing = true;
                            renderCsvPreview();
                        }
                    });
                });
                
                // Save buttons
                document.querySelectorAll('.save-csv-row-btn').forEach(btn => {
                    const index = parseInt(btn.getAttribute('data-index'));
                    btn.addEventListener('click', () => {
                        if (csvRows[index]) {
                            csvRows[index].editing = false;
                            
                            const name = csvRows[index].payload.Name;
                            const isValid = isValidName(name);
                            csvRows[index].isValid = isValid;
                            csvRows[index].include = isValid;
                            
                            // Re-check existence
                            if (isValid) {
                                findExistingByName(name).then(existingId => {
                                    csvRows[index].existingId = existingId;
                                    csvRows[index].include = !existingId;
                                    renderCsvPreview();
                                });
                            } else {
                                renderCsvPreview();
                            }
                        }
                    });
                });
                
                // Cancel buttons
                document.querySelectorAll('.cancel-csv-row-btn').forEach(btn => {
                    const index = parseInt(btn.getAttribute('data-index'));
                    btn.addEventListener('click', () => {
                        if (csvRows[index]) {
                            csvRows[index].editing = false;
                            renderCsvPreview();
                        }
                    });
                });
                
                // Edit field inputs
                document.querySelectorAll('.edit-field').forEach(input => {
                    input.addEventListener('input', () => {
                        const index = parseInt(input.getAttribute('data-index'));
                        const field = input.getAttribute('data-field');
                        if (csvRows[index]) {
                            if (field === 'OnboardingDate' || field === 'OffboardingDate') {
                                csvRows[index].payload[field] = input.value ? new Date(input.value).toISOString() : null;
                            } else {
                                csvRows[index].payload[field] = input.value;
                            }
                        }
                    });
                    
                    // Handle select changes
                    if (input.tagName === 'SELECT') {
                        input.addEventListener('change', () => {
                            const index = parseInt(input.getAttribute('data-index'));
                            const field = input.getAttribute('data-field');
                            if (csvRows[index]) {
                                csvRows[index].payload[field] = input.value;
                            }
                        });
                    }
                });
            }

            // --- IMPORT ENGINE ---
            function logCsv(msg) {
                const logEl = document.getElementById('csvLog');
                logEl.classList.remove('hidden');
                logEl.textContent += msg + "\n";
                logEl.scrollTop = logEl.scrollHeight;
            }

            async function importCsvRows() {
                const selectedRows = csvRows.filter(r => r.include);
                
                if (selectedRows.length === 0) {
                    showStatus("No rows selected for import.", "error");
                    return;
                }
                
                const validRows = selectedRows.filter(r => 
                    r.payload.Title && r.payload.Title.trim() !== ""
                );
                
                if (validRows.length === 0) {
                    showStatus("No valid rows to import (all selected rows have missing Name).", "error");
                    return;
                }
                
                if (!requestDigest) {
                    await refreshDigest();
                    if (!requestDigest) {
                        showStatus("Request digest token is missing. Please refresh the page.", "error");
                        return;
                    }
                }
                
                logCsv(`Starting import of ${validRows.length} selected rows...`);
                
                let success = 0;
                let fail = 0;
                
                for (const row of validRows) {
                    const name = row.payload.Title;
                    const i = row.index;
                    
                    try {
                        if (row.existingId) {
                            // Update existing item
                            const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${row.existingId})`, {
                                method: "POST",
                                headers: {
                                    "Accept": "application/json;odata=nometadata",
                                    "Content-Type": "application/json;odata=nometadata",
                                    "IF-MATCH": "*",
                                    "X-HTTP-Method": "MERGE",
                                    "X-RequestDigest": requestDigest
                                },
                                body: JSON.stringify(row.payload)
                            });
                            
                            if (resp.ok) {
                                logCsv(`✓ Row ${i + 1}: Updated existing record "${name}"`);
                                success++;
                            } else {
                                const errorText = await resp.text();
                                logCsv(`✗ Row ${i + 1}: ERROR updating "${name}" - ${resp.status} ${errorText}`);
                                fail++;
                            }
                        } else {
                            // Create new item
                            const resp = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items`, {
                                method: "POST",
                                headers: {
                                    "Accept": "application/json;odata=nometadata",
                                    "Content-Type": "application/json;odata=nometadata",
                                    "X-RequestDigest": requestDigest
                                },
                                body: JSON.stringify(row.payload)
                            });
                            
                            if (resp.ok) {
                                const data = await resp.json();
                                logCsv(`✓ Row ${i + 1}: Created new record "${name}" (ID ${data.Id})`);
                                success++;
                            } else {
                                const errorText = await resp.text();
                                logCsv(`✗ Row ${i + 1}: ERROR creating "${name}" - ${resp.status} ${errorText}`);
                                fail++;
                            }
                        }
                    } catch (e) {
                        console.error(`Import error for row ${i + 1}:`, e);
                        logCsv(`✗ Row ${i + 1}: SYSTEM ERROR for "${name}" - ${e.message}`);
                        fail++;
                    }
                }
                
                const importSummary = `Import complete. Success: ${success}, Failed: ${fail}.`;
                logCsv(importSummary);
                
                if (fail === 0) {
                    showStatus(importSummary, "success");
                } else {
                    showStatus(importSummary, fail === validRows.length ? "error" : "warning");
                }
                
                await loadData();
            }

            // --- MAIN VIEW RENDERING ---
            function renderMain() {
                const nameVal = document.getElementById('fName').value.toLowerCase();
                const deptVal = (document.getElementById('fDepartment') ? document.getElementById('fDepartment').value : '').toLowerCase();
                const locVal = (document.getElementById('fLocation') ? document.getElementById('fLocation').value : '').toLowerCase();
                const posVal = (document.getElementById('fPosition') ? document.getElementById('fPosition').value : '').toLowerCase();
                const monthVal = (document.getElementById('fMonth') ? document.getElementById('fMonth').value : '');
                
                const today = new Date();
                
                filteredData = masterData.filter(i => {
                    const mName = String(i.Name || "").toLowerCase().includes(nameVal);
                    const mDept = String(i.Department || "").toLowerCase().includes(deptVal);
                    const mLoc = String(i.Location || "").toLowerCase().includes(locVal);
                    const mPos = String(i.Position || "").toLowerCase().includes(posVal);

                    const itemMonth = getMonthKey(i.OnboardingDate) || getMonthKey(i.OffboardingDate);
                    const mMonth = !monthVal || itemMonth === monthVal;
                    
                    // Apply status filter
                    let mStatus = true;
                    if (selectedStatus) {
                        switch(selectedStatus) {
                            case 'employed':
                                mStatus = !i.OffboardingDate;
                                break;
                            case 'offboarded':
                                mStatus = i.OffboardingDate;
                                break;
                            case 'rehire':
                                mStatus = i.RehireFlag && i.RehireFlag.toLowerCase() === 'yes';
                                break;
                        }
                    }
                    
                    return mName && mDept && mLoc && mPos && mStatus && mMonth;
                });

                const sortFieldMap = {
                    'Name': 'Name',
                    'department': 'Department',
                    'location': 'Location',
                    'onboardingdate': 'OnboardingDate',
                    'offboardingdate': 'OffboardingDate',
                    'position': 'Position'
                };
                
                const sortField = sortFieldMap[currentSort.field] || 'Name';
                const sortedData = sortData(filteredData, sortField, currentSort.order);

                document.getElementById('mainTbody').innerHTML = sortedData.map(i => {
                    const isEmployed = !i.OffboardingDate;
                    const isOffboarded = i.OffboardingDate;
                    const rowClass = isEmployed ? 'onboarding-row' : (isOffboarded ? 'offboarding-row' : '');
                    
                    return `
                    <tr data-id="${i.Id}" class="${rowClass}">
<td>
	<input type="checkbox" class="row-check" ${selectedIds.has(i.Id) ? 'checked' : ''}>
	</td>
	<td>
		<input class="inline-input read-only" value="${escapeHtml(i.Name || '')}" readonly>
		</td>
		<td>
			<input class="inline-input read-only" value="${escapeHtml(i.Department || '')}" readonly>
			</td>
			<td>
				<input class="inline-input read-only" value="${escapeHtml(i.Location || '')}" readonly>
				</td>
				<td>
					<input type="date" class="inline-input read-only" value="${escapeHtml(formatDateForDisplay(i.OnboardingDate) || '')}" readonly>
					</td>
					<td>
						<input type="date" class="inline-input read-only" value="${escapeHtml(formatDateForDisplay(i.OffboardingDate) || '')}" readonly>
						</td>
						<td>
							<input class="inline-input read-only" value="${escapeHtml(i.OffboardingReason || '')}" readonly>
							</td>
							<td>
								<input class="inline-input read-only" value="${escapeHtml(i.Position || '')}" readonly>
								</td>
								<td>
									<select class="inline-input read-only" disabled style="background: transparent;">
										<option value="Yes" ${i.RehireFlag === 'Yes' ? 'selected' : ''}>Yes</option>
										<option value="No" ${i.RehireFlag === 'No' ? 'selected' : ''}>No</option>
										<option value="Pending" ${i.RehireFlag === 'Pending' ? 'selected' : ''}>Pending</option>
									</select>
								</td>
								<td>
									<input class="inline-input read-only" value="${escapeHtml(i.WorkEmail || '')}" readonly>
									</td>
									<td>
										<button class="btn btn-outline btn-small btn-edit">Edit</button>
										<button class="btn btn-primary btn-small btn-save hidden">Save</button>
										<button class="btn btn-warning btn-small btn-clear" title="Clear all fields except Name">Clear</button>
										<button class="btn btn-danger btn-small btn-del">Delete</button>
									</td>
								</tr>`;
                }).join('');
                
                updateStatistics(filteredData);
                updateSortUI();
                attachRowEvents();
                updateBulkActionButtons();
            }

            function attachRowEvents() {
                // Checkboxes
                document.querySelectorAll('.row-check').forEach(cb => {
                    cb.onchange = (e) => {
                        const id = parseInt(e.target.closest('tr').dataset.id);
                        e.target.checked ? selectedIds.add(id) : selectedIds.delete(id);
                        updateBulkActionButtons();
                    };
                });
                
                // Edit buttons
                document.querySelectorAll('.btn-edit').forEach(btn => {
                    btn.onclick = (e) => {
                        const tr = e.target.closest('tr');
                        tr.querySelectorAll('.inline-input, select').forEach(i => { 
                            i.readOnly = false; 
                            i.classList.remove('read-only');
                            if (i.tagName === 'SELECT') i.disabled = false;
                        });
                        tr.querySelector('.btn-edit').classList.add('hidden');
                        tr.querySelector('.btn-save').classList.remove('hidden');
                    };
                });
                
                // Save buttons
                document.querySelectorAll('.btn-save').forEach(btn => {
                    btn.onclick = async (e) => {
                        const tr = e.target.closest('tr');
                        const id = tr.dataset.id;
                        const inputs = tr.querySelectorAll('.inline-input, select');
                        
                        const name = inputs[0].value.trim();
                        
                        if (!isValidName(name)) {
                            showStatus("Name is required", "error");
                            inputs[0].style.borderColor = '#ff6666';
                            return;
                        }
                        
                        const data = {
                            Name: Name,
                            Department: inputs[1].value,
                            Location: inputs[2].value,
                            OnboardingDate: inputs[3].value ? formatDateForSharePoint(inputs[3].value) : null,
                            OffboardingDate: inputs[4].value ? formatDateForSharePoint(inputs[4].value) : null,
                            OffboardingReason: inputs[5].value,
                            Position: inputs[6].value,
                            RehireFlag: inputs[7].value,
                            WorkEmail: inputs[8].value
                        };
                        
                        if (!requestDigest) await refreshDigest();
                        
                        try {
                            await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                                method: "POST",
                                body: JSON.stringify(data),
                                headers: { 
                                    "Accept": "application/json;odata=nometadata", 
                                    "Content-Type": "application/json;odata=nometadata", 
                                    "X-RequestDigest": requestDigest, 
                                    "X-HTTP-Method": "MERGE", 
                                    "IF-MATCH": "*" 
                                }
                            });
                            showStatus("Record Updated", "success");
                            await loadData();
                        } catch (error) {
                            showStatus("Error updating record", "error");
                            console.error(error);
                        }
                    };
                });
                
                // Clear buttons
                document.querySelectorAll('.btn-clear').forEach(btn => {
                    btn.onclick = (e) => {
                        const tr = e.target.closest('tr');
                        const id = tr.dataset.id;
                        showConfirmation(id);
                    };
                });
                
                // Delete buttons
                document.querySelectorAll('.btn-del').forEach(btn => {
                    btn.onclick = async (e) => {
                        if(confirm("Delete this record?")) {
                            const id = e.target.closest('tr').dataset.id;
                            if (!requestDigest) await refreshDigest();
                            
                            try {
                                await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                                    method: "POST",
                                    headers: { 
                                        "X-RequestDigest": requestDigest, 
                                        "X-HTTP-Method": "DELETE", 
                                        "IF-MATCH": "*" 
                                    }
                                });
                                showStatus("Record Deleted", "success");
                                await loadData();
                            } catch (error) {
                                showStatus("Error deleting record", "error");
                            }
                        }
                    };
                });
            }

            function showConfirmation(recordId) {
                clearRecordId = recordId;
                document.getElementById('confirmationOverlay').classList.remove('hidden');
                document.getElementById('confirmationBox').classList.remove('hidden');
            }

            function hideConfirmation() {
                clearRecordId = null;
                document.getElementById('confirmationOverlay').classList.add('hidden');
                document.getElementById('confirmationBox').classList.add('hidden');
            }

            async function performClearRecord() {
                if (!clearRecordId) {
                    hideConfirmation();
                    return;
                }
                
                const id = clearRecordId;
                hideConfirmation();
                
                const data = { 
                    Department: "",
                    Location: "",
                    OnboardingDate: null,
                    OffboardingDate: null,
                    OffboardingReason: "",
                    Position: "",
                    RehireFlag: "",
                    WorkEmail: ""
                };
                
                if (!requestDigest) await refreshDigest();
                
                try {
                    const response = await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                        method: "POST",
                        body: JSON.stringify(data),
                        headers: { 
                            "Accept": "application/json;odata=nometadata", 
                            "Content-Type": "application/json;odata=nometadata", 
                            "X-RequestDigest": requestDigest, 
                            "X-HTTP-Method": "MERGE", 
                            "IF-MATCH": "*" 
                        }
                    });
                    
                    if (response.ok) {
                        showStatus("Record cleared successfully", "success");
                        await loadData();
                    } else {
                        const errorText = await response.text();
                        showStatus("Error clearing record", "error");
                        console.error("Clear error:", errorText);
                    }
                } catch (error) {
                    showStatus("Error clearing record", "error");
                    console.error("Clear error:", error);
                }
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
                
                updateBulkActionButtons();
            }

            function updateBulkActionButtons() {
                const deleteBtn = document.getElementById('btnDeleteSelected');
                const exportBtn = document.getElementById('btnExportSelected');
                const hasSelection = selectedIds.size > 0;
                
                deleteBtn.classList.toggle('hidden', !hasSelection);
                exportBtn.classList.toggle('hidden', !hasSelection);
                
                deleteBtn.innerText = `🗑️ Delete Selected (${selectedIds.size})`;
                exportBtn.innerText = `📥 Export Selected (${selectedIds.size})`;
            }

            async function deleteSelected() {
                if(!confirm(`Delete ${selectedIds.size} selected records? This action cannot be undone.`)) return;
                if (!requestDigest) await refreshDigest();
                
                let deleted = 0;
                let failed = 0;
                
                for(let id of selectedIds) {
                    try {
                        await fetch(`${SITE_URL}/_api/web/lists(guid'${LIST_GUID}')/items(${id})`, {
                            method: "POST", 
                            headers: { 
                                "X-RequestDigest": requestDigest, 
                                "X-HTTP-Method": "DELETE", 
                                "IF-MATCH": "*" 
                            }
                        });
                        deleted++;
                    } catch (error) {
                        failed++;
                        console.error(`Failed to delete record ${id}:`, error);
                    }
                }
                
                selectedIds.clear();
                
                if (failed === 0) {
                    showStatus(`Successfully deleted ${deleted} records`, "success");
                } else {
                    showStatus(`Deleted ${deleted} records, ${failed} failed`, "warning");
                }
                
                await loadData();
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

            init();
        })();