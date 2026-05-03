
      // CONFIG
      const siteUrl =
        "https://your-tenant.sharepoint.com/sites/IT-HelpDesk";
      const listGuid = "YOUR-NEWHIRES-LIST-GUID";
      const THEME_KEY = "nhThemePreference";

      let loadedItemId = null;
      let singleConfirmPendingFn = null;
      let requestDigest = "";
      let csvRows = []; // {index, payload, existingId, include, editing}

      // Record View State
      let allRecords = [];
      let filteredRecords = [];
      let uniquePositions = new Set();
      let uniqueDepartments = new Set();
      let uniqueHireDates = new Set();
      let selectedIds = new Set(); // IDs selected in record view

      // Search-as-you-type state
      let searchSuggestions = [];
      let activeSuggestionIndex = -1;
      let searchDebounceTimer = null;
      let filterDebounceTimer = null;

      // INTERNAL FIELD NAMES - UPDATED BASED ON XML RESPONSE
      const fields = {
        fullName: "FullName",
        firstName: "FirstName",
        lastName: "LastName",
        username: "M365Username_x002f_Computerusern",
        windowsPassword: "Windows_x002f_M365Password",
        iloan: "ILOANID",
        five9: "Five9Username",
        five9Password: "Five9Password",
        five9StationId: "Five9StationID",
        department: "Department",
        position: "Position",
        location: "Location",
        manager: "Manager",
        hireDate: "HireDate",
        phone: "PhoneNumber",
        statusUser: "UserCreated_x003f_",
        statusEmail: "EmailSent_x003f_",
        statusManager: "ManagerAssigned_x003f_",
        stationAssigned: "StationAssigned_x003f_",
        unifi: "Unifi",
        profileCreated: "ProfileCreated_x003f_",
        fusionId: "Fusionid_x0028_Philliphinesonly_",
      };

      // ==================== ENHANCED CSV HEADER NORMALIZATION ====================
      function normalizeHeader(h) {
        if (!h) return "";
        
        const normalized = h.toLowerCase().replace(/[^a-z0-9]/g, "");
        
        // Define comprehensive header mappings
        const headerMappings = {
          // Full Name mappings - handles "Name", "FULL NAME", "Full Name", etc.
          "fullname": ["fullname", "name", "fullname", "fullname", "fullname"],
          
          // FusionID mappings - handles "External Email", "BPO Email", "Fusion Email", etc.
          "fusionid": ["fusionid", "bpoemail", "externalemail", "fusionemail", "bpoemail", "bpomail", "fusionidRegion-PHonly"],
          
          // First Name mappings
          "firstname": ["firstname", "firstname", "firstname"],
          
          // Last Name mappings
          "lastname": ["lastname", "lastname", "lastname"],
          
          // Phone Number mappings - handles "Standardized Number"
          "phonenumber": ["phonenumber", "phonenumber", "standardizednumber", "phonenumber"],
          
          // Position/Job Title mappings
          "position": ["position", "jobtitle", "jobtitle", "position", "title"],
          
          // M365 Username mappings - handles various formats
          "m365usernamecomputerusername": ["m365usernamecomputerusername", "m365usernamecomputerusername", "m365username", "computerusername", "username"],
          
          // Windows Password mappings
          "windowsm365password": ["windowsm365password", "windowsm365password", "m365password", "windowsm365password"],
          
          // Department mappings
          "department": ["department", "department"],
          
          // Location mappings
          "location": ["location", "location"],
          
          // Manager mappings
          "manager": ["manager", "manager"],
          
          // Hire Date mappings
          "hiredate": ["hiredate", "hiredate", "hiredate"],
          
          // Five9 Username mappings
          "five9username": ["five9username", "five9username"],
          
          // Five9 Password mappings
          "five9password": ["five9password", "five9password"],
          
          // Five9 Station ID mappings
          "five9stationid": ["five9stationid", "five9stationid", "stationid"],
          
          // ILOAN ID mappings
          "iloanid": ["iloanid", "iloanid", "iloan"]
        };
        
        // Find the matching header
        for (const [targetHeader, variations] of Object.entries(headerMappings)) {
          if (variations.includes(normalized)) {
            return targetHeader;
          }
        }
        
        // If no match found, return the normalized version
        return normalized;
      }

      // ==================== NEW DATE PARSING FUNCTIONS ====================
      function parseDateInput(dateStr) {
        if (!dateStr || dateStr.trim() === '') return null;
        
        try {
          // Remove any whitespace
          const cleaned = dateStr.trim();
          
          // Handle various formats
          let date;
          
          // Try parsing as ISO format first (yyyy-MM-dd)
          if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
            date = new Date(cleaned);
          } 
          // Handle formats like 1/5/2026, 01/05/2026
          else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cleaned)) {
            const parts = cleaned.split('/');
            const month = parseInt(parts[0], 10) - 1;
            const day = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
          // Handle formats like 1-5-2026
          else if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(cleaned)) {
            const parts = cleaned.split('-');
            const month = parseInt(parts[0], 10) - 1;
            const day = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
          // Handle formats like 2026/1/5
          else if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(cleaned)) {
            const parts = cleaned.split('/');
            const year = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const day = parseInt(parts[2], 10);
            date = new Date(year, month, day);
          }
          // Handle any other format that Date can parse
          else {
            date = new Date(cleaned);
          }
          
          // Check if date is valid
          if (isNaN(date.getTime())) {
            console.warn(`Invalid date format: ${dateStr}`);
            return null;
          }
          
          return date;
        } catch (e) {
          console.error('Error parsing date:', e);
          return null;
        }
      }

      function formatDateForSharePoint(dateObj) {
        if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
          return null;
        }
        
        // Use toISOString to get the correct UTC representation for local date
        return dateObj.toISOString();
      }

      function formatDateForInput(dateStr) {
        if (!dateStr) return '';
        
        const date = parseDateInput(dateStr);
        if (!date) return '';
        
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        
        return `${year}-${month}-${day}`;
      }

      function formatDateForDisplay(dateStr) {
        if (!dateStr) return '';
        
        const date = parseDateInput(dateStr);
        if (!date) return dateStr;
        
        const month = date.getMonth() + 1;
        const day = date.getDate();
        const year = date.getFullYear();
        
        return `${month}/${day}/${year}`;
      }

      // ==================== PHONE NUMBER FORMATTER ====================
      function formatPhoneNumber(phone) {
        if (!phone || phone.trim() === '') return '';
        
        let cleaned = phone.trim();
        
        // Remove all non-digit characters except +
        cleaned = cleaned.replace(/[^\d+]/g, '');
        
        // If it doesn't start with +, add it
        if (!cleaned.startsWith('+')) {
          // If it starts with 1 (US country code), add +1
          if (cleaned.startsWith('1') && cleaned.length >= 10) {
            cleaned = '+' + cleaned;
          }
          // Otherwise, add + for international format
          else {
            cleaned = '+' + cleaned;
          }
        }
        
        return cleaned;
      }

      // ==================== HELPER FUNCTIONS ====================
      function escapeHtml(str) {
        return String(str).replace(/[&<>\"']/g, (c) => {
          return (
            {
              "&": "&amp;",
              "<": "&lt;",
              ">": "&gt;",
              '"': "&quot;",
              "'": "&#39;",
            }[c] || c
          );
        });
      }

      // ==================== THEME HANDLING ====================
      function applyTheme(theme) {
        const docEl = document.documentElement;
        docEl.setAttribute("data-theme", theme);
        const btn = document.getElementById("themeToggle");
        if (!btn) return;
        const icon = btn.querySelector(".theme-toggle-icon");
        const text = btn.querySelector(".theme-toggle-text");

        if (theme === "dark") {
          if (icon) icon.textContent = "☀";
          if (text) text.textContent = "Light mode";
        } else {
          if (icon) icon.textContent = "🌙";
          if (text) text.textContent = "Dark mode";
        }
      }

      function initTheme() {
        try {
          const saved = localStorage.getItem(THEME_KEY);
          const initial = saved === "dark" ? "dark" : "light";
          applyTheme(initial);
        } catch (e) {
          applyTheme("light");
        }

        const btn = document.getElementById("themeToggle");
        if (btn) {
          btn.addEventListener("click", () => {
            const current =
              document.documentElement.getAttribute("data-theme") || "light";
            const next = current === "dark" ? "light" : "dark";
            applyTheme(next);
            try {
              localStorage.setItem(THEME_KEY, next);
            } catch (e) {}
          });
        }
      }

      // ==================== TABS ====================
      function showTab(tab) {
        document.getElementById("tab-single").style.display =
          tab === "single" ? "block" : "none";
        document.getElementById("tab-csv").style.display =
          tab === "csv" ? "block" : "none";
        document.getElementById("tab-record").style.display =
          tab === "record" ? "block" : "none";

        document.getElementById("btn-tab-single").classList.toggle("active", tab === "single");
        document.getElementById("btn-tab-csv").classList.toggle("active", tab === "csv");
        document.getElementById("btn-tab-record").classList.toggle("active", tab === "record");

        if (tab === 'record') {
          loadAllRecords();
        }

        if (document.getElementById("statusBox").textContent) {
          document.getElementById("statusBox").style.display = "block";
        }
      }

      // ==================== DIGEST ====================
      async function refreshDigest() {
        try {
          const resp = await fetch(`${siteUrl}/_api/contextinfo`, {
            method: "POST",
            headers: { Accept: "application/json;odata=nometadata" },
          });

          if (!resp.ok) {
            console.error(
              "Failed to fetch request digest. Status:",
              resp.status,
              "StatusText:",
              resp.statusText
            );
            return;
          }

          const data = await resp.json();
          requestDigest = data.FormDigestValue;
          console.log("Request digest refreshed successfully.");
          setTimeout(refreshDigest, 1500000);
        } catch (e) {
          console.error(
            "Error refreshing request digest (Check network connection or siteUrl):",
            e
          );
          setTimeout(refreshDigest, 300000);
        }
      }
      refreshDigest();

      // ==================== RECORD VIEW FUNCTIONS ====================
      async function loadAllRecords() {
        const recordsBody = document.getElementById("recordsTableBody");
        const recordCount = document.getElementById("recordCount");

        selectedIds.clear();
        updateActionBar();
        recordsBody.innerHTML = `<tr><td colspan="21" style="text-align:center;padding:30px;"><div class="loading-spinner"></div> Loading records...</td></tr>`;

        try {
          const selectFields = ["Id"].concat(Object.values(fields));
          const endpoint = `${siteUrl}/_api/web/lists(guid'${listGuid}')/items?$top=5000&$select=${selectFields.join(",")}&$orderby=Id desc`;

          const resp = await fetch(endpoint, {
            headers: { Accept: "application/json;odata=nometadata" },
          });

          if (!resp.ok) {
            throw new Error(`Failed to load records: ${resp.status}`);
          }

          const data = await resp.json();
          allRecords = data.value || [];

          extractUniqueValues();
          updateItStats();
          applyFilters();

        } catch (e) {
          console.error("Error loading records:", e);
          recordsBody.innerHTML = `<tr><td colspan="21" style="text-align:center;padding:30px;color:var(--fluent-danger);">Error loading records: ${e.message}</td></tr>`;
          recordCount.textContent = "Error loading records";
        }
      }

      function extractUniqueValues() {
        uniquePositions.clear();
        uniqueDepartments.clear();
        uniqueHireDates.clear();
        
        allRecords.forEach(record => {
          if (record[fields.position]) {
            uniquePositions.add(record[fields.position]);
          }
          if (record[fields.department]) {
            uniqueDepartments.add(record[fields.department]);
          }
          if (record[fields.hireDate]) {
            const formattedDate = formatDateForDisplay(record[fields.hireDate]);
            if (formattedDate) {
              uniqueHireDates.add(formattedDate);
            }
          }
        });
        
        // Sort and populate dropdowns
        populateFilterDropdowns();
      }

      function populateFilterDropdowns() {
        const positionFilter = document.getElementById("filterPosition");
        const departmentFilter = document.getElementById("filterDepartment");
        const hireDateFilter = document.getElementById("filterHireDate");
        
        // Save current selections
        const currentPosition = positionFilter.value;
        const currentDepartment = departmentFilter.value;
        const currentHireDate = hireDateFilter.value;
        
        // Populate Position dropdown
        positionFilter.innerHTML = '<option value="">All Positions</option>';
        Array.from(uniquePositions).sort().forEach(pos => {
          positionFilter.innerHTML += `<option value="${escapeHtml(pos)}">${escapeHtml(pos)}</option>`;
        });
        
        // Populate Department dropdown
        departmentFilter.innerHTML = '<option value="">All Departments</option>';
        Array.from(uniqueDepartments).sort().forEach(dept => {
          departmentFilter.innerHTML += `<option value="${escapeHtml(dept)}">${escapeHtml(dept)}</option>`;
        });
        
        // Populate Hire Date dropdown
        hireDateFilter.innerHTML = '<option value="">All Dates</option>';
        Array.from(uniqueHireDates).sort().reverse().forEach(date => {
          hireDateFilter.innerHTML += `<option value="${escapeHtml(date)}">${escapeHtml(date)}</option>`;
        });
        
        // Restore selections if they still exist
        if (currentPosition && uniquePositions.has(currentPosition)) {
          positionFilter.value = currentPosition;
        }
        if (currentDepartment && uniqueDepartments.has(currentDepartment)) {
          departmentFilter.value = currentDepartment;
        }
        if (currentHireDate && uniqueHireDates.has(currentHireDate)) {
          hireDateFilter.value = currentHireDate;
        }
      }

      function applyFilters() {
        const nameFilter = document.getElementById("filterName").value.toLowerCase().trim();
        const positionFilter = document.getElementById("filterPosition").value;
        const departmentFilter = document.getElementById("filterDepartment").value;
        const hireDateFilter = document.getElementById("filterHireDate").value;
        
        filteredRecords = allRecords.filter(record => {
          // Name filter
          if (nameFilter) {
            const fullName = (record[fields.fullName] || "").toLowerCase();
            if (!fullName.includes(nameFilter)) {
              return false;
            }
          }
          
          // Position filter
          if (positionFilter) {
            if (record[fields.position] !== positionFilter) {
              return false;
            }
          }
          
          // Department filter
          if (departmentFilter) {
            if (record[fields.department] !== departmentFilter) {
              return false;
            }
          }
          
          // Hire Date filter
          if (hireDateFilter) {
            const formattedDate = formatDateForDisplay(record[fields.hireDate]);
            if (formattedDate !== hireDateFilter) {
              return false;
            }
          }
          
          return true;
        });
        
        renderRecordsTable();
        updateRecordCount();
      }


      function itBadge(value) {
        const yes = value && (value.toUpperCase() === 'YES' || value === 'Yes');
        return yes
          ? '<span class="it-badge it-badge-yes">YES</span>'
          : '<span class="it-badge it-badge-no">NO</span>';
      }

      function renderRecordsTable() {
        const recordsBody = document.getElementById("recordsTableBody");

        if (filteredRecords.length === 0) {
          recordsBody.innerHTML = '<tr><td colspan="21" style="text-align:center;padding:30px;">No records found matching the filters.</td></tr>';
          updateActionBar();
          return;
        }

        let html = '';
        filteredRecords.forEach(record => {
          const id = record.Id;
          const checked = selectedIds.has(id) ? 'checked' : '';
          const fullName       = record[fields.fullName]        || '';
          const username       = record[fields.username]        || '';
          const winPwd         = record[fields.windowsPassword] || '';
          const five9User      = record[fields.five9]           || '';
          const five9Pwd       = record[fields.five9Password]   || '';
          const stationId      = record[fields.five9StationId]  || '';
          const iloan          = record[fields.iloan]           || '';
          const fusionId       = record[fields.fusionId]        || '';
          const dept           = record[fields.department]      || '';
          const position       = record[fields.position]        || '';
          const hireDate       = formatDateForDisplay(record[fields.hireDate]);
          const location       = record[fields.location]        || '';
          const manager        = record[fields.manager]         || '';
          const phone          = record[fields.phone]           || '';
          const profileCreated = record[fields.profileCreated]  || '';
          const unifi          = record[fields.unifi]           || '';
          const statusManager  = record[fields.statusManager]   || '';
          const statusEmail    = record[fields.statusEmail]     || '';
          const stationAssigned= record[fields.stationAssigned] || '';

          html += `<tr data-id="${id}">
            <td style="text-align:center;"><input type="checkbox" class="row-check" ${checked} data-id="${id}"></td>
            <td><strong>${escapeHtml(fullName)}</strong></td>
            <td>${escapeHtml(username)}</td>
            <td>${escapeHtml(hireDate)}</td>
            <td>${escapeHtml(dept)}</td>
            <td>${escapeHtml(position)}</td>
            <td>${escapeHtml(location)}</td>
            <td>${escapeHtml(manager)}</td>
            <td>${escapeHtml(winPwd)}</td>
            <td>${escapeHtml(five9User)}</td>
            <td>${escapeHtml(five9Pwd)}</td>
            <td>${escapeHtml(stationId)}</td>
            <td>${escapeHtml(iloan)}</td>
            <td>${escapeHtml(fusionId)}</td>
            <td>${escapeHtml(phone)}</td>
            <td>${itBadge(profileCreated)}</td>
            <td>${itBadge(unifi)}</td>
            <td>${itBadge(statusManager)}</td>
            <td>${itBadge(statusEmail)}</td>
            <td>${itBadge(stationAssigned)}</td>
            <td><span class="edit-link" data-edit-id="${id}">Edit</span></td>
          </tr>`;
        });

        recordsBody.innerHTML = html;
        attachRecordCheckboxEvents();
        updateSelectAllCheckbox();
      }

      function updateRecordCount() {
        const recordCount = document.getElementById("recordCount");
        recordCount.textContent = `Showing ${filteredRecords.length} of ${allRecords.length} records`;
      }

      window.loadRecordForEdit = function(id) {
        const record = allRecords.find(r => r.Id === id);
        if (record) {
          loadedItemId = id;
          populateForm(record);
          updateSingleTabButtons();
          showTab('single');
          showStatus(`Record loaded for editing (ID: ${id}). Update fields then click "Update Existing".`, 'success');
          const container = document.querySelector('.nh-container');
          if (container) container.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      };

      function attachRecordCheckboxEvents() {
        document.querySelectorAll('.row-check').forEach(cb => {
          cb.addEventListener('change', () => {
            const id = parseInt(cb.getAttribute('data-id'));
            if (cb.checked) {
              selectedIds.add(id);
            } else {
              selectedIds.delete(id);
            }
            updateActionBar();
            updateSelectAllCheckbox();
          });
        });

        document.querySelectorAll('.edit-link[data-edit-id]').forEach(link => {
          link.addEventListener('click', () => {
            const id = parseInt(link.getAttribute('data-edit-id'));
            loadRecordForEdit(id);
          });
        });
      }

      function updateSelectAllCheckbox() {
        const sa = document.getElementById('selectAllRecords');
        if (!sa) return;
        const boxes = document.querySelectorAll('.row-check');
        if (boxes.length === 0) { sa.checked = false; sa.indeterminate = false; return; }
        const checkedCount = [...boxes].filter(b => b.checked).length;
        if (checkedCount === 0)            { sa.checked = false; sa.indeterminate = false; }
        else if (checkedCount === boxes.length) { sa.checked = true;  sa.indeterminate = false; }
        else                               { sa.checked = false; sa.indeterminate = true; }
      }

      function toggleAllRecords(e) {
        const checked = e.target.checked;
        selectedIds.clear();
        document.querySelectorAll('.row-check').forEach(cb => {
          cb.checked = checked;
          if (checked) {
            const id = parseInt(cb.getAttribute('data-id'));
            selectedIds.add(id);
          }
        });
        updateActionBar();
      }

      function updateActionBar() {
        const bar   = document.getElementById('recordActionBar');
        const count = document.getElementById('selectionCount');
        if (!bar) return;
        if (selectedIds.size === 0) {
          bar.style.display = 'none';
        } else {
          bar.style.display = 'flex';
          if (count) count.textContent = `${selectedIds.size} selected`;
        }
      }

      function updateItStats() {
        const total = allRecords.length;
        const totalEl       = document.getElementById('stat-total');
        const monthEl       = document.getElementById('stat-month');
        const weekEl        = document.getElementById('stat-week');
        const pendingM365El = document.getElementById('stat-pending-m365');
        const pendingEmailEl= document.getElementById('stat-pending-email');

        if (totalEl) totalEl.textContent = total || '—';
        if (total === 0) return;

        const now = new Date();
        const thisMonth = allRecords.filter(r => {
          if (!r[fields.hireDate]) return false;
          const d = new Date(r[fields.hireDate]);
          return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        }).length;
        if (monthEl) monthEl.textContent = `this month: ${thisMonth}`;

        const weekStart = new Date(now);
        weekStart.setDate(now.getDate() - now.getDay());
        weekStart.setHours(0, 0, 0, 0);
        const thisWeek = allRecords.filter(r => {
          if (!r[fields.hireDate]) return false;
          const d = new Date(r[fields.hireDate]);
          return d >= weekStart;
        }).length;
        if (weekEl) weekEl.textContent = thisWeek;

        const isYes = v => v && v.toUpperCase() === 'YES';
        const pendingM365  = allRecords.filter(r => !isYes(r[fields.profileCreated])).length;
        const pendingEmail = allRecords.filter(r => !isYes(r[fields.statusEmail])).length;
        if (pendingM365El)  pendingM365El.textContent  = pendingM365;
        if (pendingEmailEl) pendingEmailEl.textContent = pendingEmail;
      }


      // ==================== EXPORT RECORDS CSV ====================
      function buildCsvContent(records) {
        const headers = [
          'Full Name', 'First Name', 'Last Name',
          'M365 Username', 'M365 Password',
          'Five9 Username', 'Five9 Station ID',
          'ILOAN ID', 'Department', 'Position', 'Hire Date',
          'Location', 'Manager', 'External Email'
        ];
        const rows = records.map(r => [
          r[fields.fullName]        || '',
          r[fields.firstName]       || '',
          r[fields.lastName]        || '',
          r[fields.username]        || '',
          r[fields.windowsPassword] || '',
          r[fields.five9]           || '',
          r[fields.five9StationId]  || '',
          r[fields.iloan]           || '',
          r[fields.department]      || '',
          r[fields.position]        || '',
          formatDateForDisplay(r[fields.hireDate]) || '',
          r[fields.location]        || '',
          r[fields.manager]         || '',
          r[fields.fusionId]        || '',
        ]);
        const esc = v => `"${String(v).replace(/"/g, '""')}"`;
        return [headers.map(esc).join(',')]
          .concat(rows.map(row => row.map(esc).join(',')))
          .join('\r\n');
      }

      function exportRecordsCsv() {
        const records = selectedIds.size > 0
          ? allRecords.filter(r => selectedIds.has(r.Id))
          : (filteredRecords.length > 0 ? filteredRecords : allRecords);
        if (records.length === 0) { showStatus('No records to export.', 'error'); return; }

        const csv = buildCsvContent(records);
        const today = new Date().toISOString().split('T')[0];
        const filename = `IT_NewHires_${today}.csv`;
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
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
        const scope = selectedIds.size > 0 ? `${records.length} selected record(s)` : `${records.length} record(s)`;
        showStatus(`Exported ${scope} to ${filename}.`, 'success');
      }

      function downloadLMProfileCSV() {
        const records = allRecords.filter(r => selectedIds.has(r.Id));
        if (records.length === 0) { showStatus('Please select records to download LM Profile CSV.', 'error'); return; }

        const esc = v => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
        const headers = ['First Name', 'Last Name', 'Email', 'PIMS Name', 'iLoan User ID', 'Five9 Name', 'Hire Date'];
        const rows = records.map(r => [
          r[fields.firstName]  || '',
          r[fields.lastName]   || '',
          r[fields.username]   || '',
          '',
          r[fields.iloan]      || '',
          r[fields.five9]      || '',
          formatDateForDisplay(r[fields.hireDate]) || '',
        ]);
        const csv = [headers.map(esc).join(',')]
          .concat(rows.map(row => row.map(esc).join(',')))
          .join('\r\n');

        const today = new Date().toISOString().split('T')[0];
        const filename = `LM_Profile_${today}_${records.length}_records.csv`;
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
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
        showStatus(`LM Profile CSV downloaded (${records.length} records).`, 'success');
      }

      // ==================== IT EMAIL ====================
      let itEmailCsvBlobUrl = null;

      function showItEmailModal() {
        if (selectedIds.size === 0) return;
        const modal   = document.getElementById('itEmailModal');
        const overlay = document.getElementById('itEmailModalOverlay');
        if (!modal || !overlay) return;

        const selected = allRecords.filter(r => selectedIds.has(r.Id));
        const first    = selected[0];
        const dept     = first[fields.department] || 'Various';
        const location = first[fields.location]   || 'Various';

        let earliestDate = null;
        selected.forEach(r => {
          if (r[fields.hireDate]) {
            const d = new Date(r[fields.hireDate]);
            if (!earliestDate || d < earliestDate) earliestDate = d;
          }
        });
        const hireDateStr = earliestDate ? formatDateForDisplay(earliestDate.toISOString()) : 'Various';

        document.getElementById('itEmailTo').value      = '';
        document.getElementById('itEmailSubject').value = `New Hire IT Setup - ${dept} - ${location} - ${hireDateStr}`;
        document.getElementById('itEmailBody').value    = buildItEmailBody(selected, dept, location, hireDateStr);

        if (itEmailCsvBlobUrl) URL.revokeObjectURL(itEmailCsvBlobUrl);
        const csv  = buildCsvContent(selected);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        itEmailCsvBlobUrl = URL.createObjectURL(blob);
        const today    = new Date().toISOString().split('T')[0];
        const filename = `IT_NewHires_${today}_${selected.length}_records.csv`;
        const container = document.getElementById('itEmailCsvContainer');
        if (container) {
          container.innerHTML = `
            <a href="${itEmailCsvBlobUrl}" download="${filename}"
               style="display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;
                      color:var(--fluent-primary);text-decoration:none;padding:5px 10px;
                      border:1px solid var(--fluent-primary);border-radius:var(--fluent-radius-md);
                      background:var(--fluent-primary-soft);">
              &#x1F4E5; Download Credential Sheet (${selected.length} records)
            </a>
            <div style="font-size:11px;color:var(--fluent-text-subtle);margin-top:5px;">
              Download and attach this file to your email.
            </div>`;
        }

        overlay.classList.add('visible');
        modal.classList.add('visible');
      }

      function hideItEmailModal() {
        document.getElementById('itEmailModalOverlay').classList.remove('visible');
        document.getElementById('itEmailModal').classList.remove('visible');
      }

      function buildItEmailBody(records, dept, location, hireDateStr) {
        const names = records.map(r => r[fields.fullName] || '(unknown)').join(', ');
        let body  = `Hi,\n\n`;
        body += `The IT credential sheet has been prepared for ${records.length} new hire(s) starting on ${hireDateStr} at ${location}.\n\n`;
        body += `Please review the attached credential file for full details.\n\n`;
        body += `Thank you.\n\n`;
        body += `---\n`;
        body += `Summary:\n`;
        body += `• Department: ${dept}\n`;
        body += `• Location:   ${location}\n`;
        body += `• Start Date: ${hireDateStr}\n`;
        body += records.length <= 5
          ? `• Names:      ${names}\n`
          : `• Count:      ${records.length} hires\n`;
        body += `\nRegards,\nIT Team`;
        return body;
      }

      function openItEmailClient() {
        const to      = document.getElementById('itEmailTo').value.trim();
        const subject = document.getElementById('itEmailSubject').value.trim();
        const body    = document.getElementById('itEmailBody').value;
        window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        hideItEmailModal();
        showStatus('Email client opened.', 'success');
      }

      // ==================== CSV TEMPLATE FUNCTIONS ====================
      function downloadTemplate(type) {
        const today = new Date().toISOString().split('T')[0];
        
        let csvContent = '';
        let filename = '';
        
        if (type === 'csr') {
          filename = `CSR_Template_${today}.csv`;
          csvContent = `Full name,First name,Last name,M365 Username/Computer username,Windows/ M365 Password,Five9 Username,Five9 Password,Five9 Station ID,ILOAN ID,Department,Position,Hire Date,Location,Manager,Phone number,FusionID (Region-specific field)
"John Doe","John","Doe","john.doe@company.com","Light22@@","john.doe","Light22@@","CSR001","ILOAN123456","Customer Service","Customer Service Representative","${today}","Main Office","Jane Smith","+1 (555) 123-4567","BPO-JohnDoe"
"Jane Smith","Jane","Smith","jane.smith@company.com","Light22@@","jane.smith","Light22@@","CSR002","ILOAN789012","Customer Service","Customer Service Representative","${today}","Main Office","John Manager","+1 (555) 987-6543","BPO-JaneSmith"`;
          
          showStatus(`CSR template downloaded. Department: Customer Service, Title: Customer Service Representative`, "success");
        } else if (type === 'assessment') {
          filename = `Assessment_Template_${today}.csv`;
          csvContent = `Full name,First name,Last name,M365 Username/Computer username,Windows/ M365 Password,Five9 Username,Five9 Password,Five9 Station ID,ILOAN ID,Department,Position,Hire Date,Location,Manager,Phone number,FusionID (Region-specific field)
"Alex Johnson","Alex","Johnson","alex.johnson@company.com","Light22@@","alex.johnson","Light22@@","ASSESS001","ILOAN345678","Assessment","Assessor","${today}","Assessment Center","Michael Brown","+1 (555) 234-5678","BPO-AlexJohnson"
"Sarah Williams","Sarah","Williams","sarah.williams@company.com","Light22@@","sarah.williams","Light22@@","ASSESS002","ILOAN901234","Assessment","Assessor","${today}","Assessment Center","Michael Brown","+1 (555) 876-5432","BPO-SarahWilliams"`;
          
          showStatus(`Assessment template downloaded. Department: Assessment, Title: Assessor`, "success");
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
      }

      // ==================== INITIALIZATION ====================
      initForm();
      initTheme();
      initSearchAutocomplete();
      initEventListeners();
      initRecordViewEvents();
      updateSingleTabButtons();

      async function initForm() {
        await loadChoiceColumns();
      }

      function initEventListeners() {
        // Tab buttons
        document.getElementById('btn-tab-single').addEventListener('click', () => showTab('single'));
        document.getElementById('btn-tab-csv').addEventListener('click', () => showTab('csv'));
        document.getElementById('btn-tab-record').addEventListener('click', () => showTab('record'));
        
        // Single entry buttons
        document.getElementById('loadUserBtn').addEventListener('click', loadExistingUser);
        document.getElementById('saveNewHireBtn').addEventListener('click', () => {
          const name = document.getElementById('f_fullName').value.trim();
          showSingleConfirm(`Save "${name || 'new hire'}" as a new record?`, submitNewHire);
        });
        document.getElementById('updateHireBtn').addEventListener('click', () => {
          const name = document.getElementById('f_fullName').value.trim();
          showSingleConfirm(`Update existing record for "${name || 'this user'}"?`, updateHire);
        });
        document.getElementById('btnSingleConfirmCancel').addEventListener('click', hideSingleConfirm);
        document.getElementById('btnSingleConfirmOk').addEventListener('click', executeSingleConfirm);
        document.getElementById('singleConfirmOverlay').addEventListener('click', hideSingleConfirm);
        document.getElementById('searchClearBtn').addEventListener('click', clearSearch);
        
        // CSV buttons
        document.getElementById('csrTemplateBtn').addEventListener('click', () => downloadTemplate('csr'));
        document.getElementById('assessmentTemplateBtn').addEventListener('click', () => downloadTemplate('assessment'));
        document.getElementById('csvFile').addEventListener('change', handleCsvFile);
        document.getElementById('importCsvBtn').addEventListener('click', importCsvRows);
        
        // Record View filters
        document.getElementById('filterName').addEventListener('input', () => {
          if (filterDebounceTimer) clearTimeout(filterDebounceTimer);
          filterDebounceTimer = setTimeout(applyFilters, 300);
        });
        
        document.getElementById('filterPosition').addEventListener('change', applyFilters);
        document.getElementById('filterDepartment').addEventListener('change', applyFilters);
        document.getElementById('filterHireDate').addEventListener('change', applyFilters);
        document.getElementById('refreshRecordsBtn').addEventListener('click', loadAllRecords);
        
        // Date input formatting on blur
        document.getElementById('f_hiredate').addEventListener('blur', function(e) {
          const value = e.target.value.trim();
          if (value) {
            const parsed = parseDateInput(value);
            if (parsed) {
              e.target.value = formatDateForInput(value);
            } else {
              showStatus('Invalid date format. Please use MM/DD/YYYY or YYYY-MM-DD', 'error');
            }
          }
        });
        
        // Phone input formatting on blur
        document.getElementById('f_phone').addEventListener('blur', function(e) {
          const value = e.target.value.trim();
          if (value) {
            const formatted = formatPhoneNumber(value);
            e.target.value = formatted;
          }
        });
        
        // FusionID input formatting on blur
        document.getElementById('f_fusionId').addEventListener('blur', function(e) {
          const value = e.target.value.trim();
          // Optional: Add any specific formatting for FusionID here
        });
      }

      // ==================== RECORD VIEW & MODAL EVENT WIRING ====================
      function initRecordViewEvents() {
        const safeOn = (id, fn) => {
          const el = document.getElementById(id);
          if (el) el.addEventListener('click', fn);
        };
        safeOn('selectAllRecords',  toggleAllRecords);
        safeOn('btnM365Profile',    markM365Profile);
        safeOn('btnUnifiProfile',   markUnifiProfile);
        safeOn('btnAssignManager',  markAssignManager);
        safeOn('btnSendEmail',      markSendEmail);
        safeOn('btnAssignStation',  markAssignStation);
        safeOn('btnDeleteSelected', showDeleteModal);
        safeOn('btnCancelDelete',   hideDeleteModal);
        safeOn('btnConfirmDelete',  performDelete);
        safeOn('exportCsvBtn',      exportRecordsCsv);
        safeOn('btnLMProfileCSV',   downloadLMProfileCSV);
        safeOn('btnItEmail',        showItEmailModal);
        safeOn('btnItEmailCancel',  hideItEmailModal);
        safeOn('btnItEmailOpen',    openItEmailClient);
        const itOverlay = document.getElementById('itEmailModalOverlay');
        if (itOverlay) itOverlay.addEventListener('click', hideItEmailModal);
        const overlay = document.getElementById('deleteModalOverlay');
        if (overlay) overlay.addEventListener('click', hideDeleteModal);
      }

      // ==================== CHOICE COLUMNS LOADING ====================
      async function loadChoiceColumns() {
        const endpoint = `${siteUrl}/_api/web/lists(guid'${listGuid}')/fields?$select=InternalName,Choices`;
        try {
          const resp = await fetch(endpoint, {
            headers: { Accept: "application/json;odata=nometadata" },
          });
          if (!resp.ok) {
            console.error(
              "Failed to load choice columns. Status:",
              resp.status,
              "StatusText:",
              resp.statusText,
              "Endpoint:",
              endpoint
            );
            return;
          }
          const data = await resp.json();

          data.value.forEach((f) => {
            if (!f.Choices) return;

            let dropdown = null;

            switch (f.InternalName) {
              case fields.statusUser:
                dropdown = "f_status_user";
                break;
              case fields.statusEmail:
                dropdown = "f_status_email";
                break;
              case fields.statusManager:
                dropdown = "f_status_manager";
                break;
              case fields.stationAssigned:
                dropdown = "f_station_assigned";
                break;
              case fields.unifi:
                dropdown = "f_unifi";
                break;
            }

            if (dropdown) {
              const sel = document.getElementById(dropdown);
              sel.innerHTML = `<option value="">-- Select Status --</option>`;
              f.Choices.forEach(
                (c) => (sel.innerHTML += `<option value="${c}">${c}</option>`)
              );
            }
          });
          console.log("Choice columns loaded successfully.");
        } catch (e) {
          console.error(
            "Error loading choice columns (Check network connection or siteUrl/listGuid):",
            e
          );
        }
      }

      // ==================== SEARCH EXISTING ====================
      async function loadExistingUser() {
        const searchRaw = document.getElementById("nh-search").value.trim();
        if (!searchRaw)
          return showStatus("Enter search text to load a user.", "error");

        const search = searchRaw.replace(/'/g, "''");

        const filter = `$filter=substringof('${search}', ${fields.username}) or substringof('${search}', ${fields.fullName}) or substringof('${search}', ${fields.iloan}) or substringof('${search}', ${fields.five9}) or substringof('${search}', ${fields.fusionId})`;

        const selectFields = ["Id"].concat(Object.values(fields));
        const endpoint = `${siteUrl}/_api/web/lists(guid'${listGuid}')/items?${filter}&$top=1&$select=${selectFields.join(
          ","
        )}`;

        try {
          const resp = await fetch(endpoint, {
            headers: { Accept: "application/json;odata=nometadata" },
          });
          if (!resp.ok)
            return showStatus("Search failed. HTTP " + resp.status, "error");

          const data = await resp.json();
          if (!data.value || data.value.length === 0) {
            return showStatus(
              "No user found matching the criteria.",
              "error"
            );
          }

          loadedItemId = data.value[0].Id;
          populateForm(data.value[0]);
          updateSingleTabButtons();

          showStatus(
            "User loaded successfully (ID: " + loadedItemId + "). You can now update their details.",
            "success"
          );
        } catch (e) {
          console.error(e);
          showStatus("Unexpected error while loading user.", "error");
        }
      }

      // ==================== FORM POPULATION ====================
      function populateForm(u) {
        document.getElementById("f_fullName").value =
          u[fields.fullName] || "";
        document.getElementById("f_username").value =
          u[fields.username] || "";
        document.getElementById("f_firstName").value =
          u[fields.firstName] || "";
        document.getElementById("f_lastName").value =
          u[fields.lastName] || "";
        document.getElementById("f_winPwd").value =
          u[fields.windowsPassword] || "";
        document.getElementById("f_iloan").value = u[fields.iloan] || "";
        document.getElementById("f_five9").value = u[fields.five9] || "";
        document.getElementById("f_five9Pwd").value =
          u[fields.five9Password] || "";
        document.getElementById("f_five9Station").value =
          u[fields.five9StationId] || "";
        document.getElementById("f_phone").value = u[fields.phone] || "";
        document.getElementById("f_department").value =
          u[fields.department] || "";
        document.getElementById("f_position").value =
          u[fields.position] || "";
        document.getElementById("f_location").value =
          u[fields.location] || "";
        document.getElementById("f_manager").value = u[fields.manager] || "";
        document.getElementById("f_fusionId").value = u[fields.fusionId] || "";

        if (u[fields.hireDate]) {
          const formattedDate = formatDateForInput(u[fields.hireDate]);
          document.getElementById("f_hiredate").value = formattedDate;
        } else {
          document.getElementById("f_hiredate").value = "";
        }

        document.getElementById("f_status_user").value =
          u[fields.statusUser] || "";
        document.getElementById("f_status_email").value =
          u[fields.statusEmail] || "";
        document.getElementById("f_status_manager").value =
          u[fields.statusManager] || "";
        document.getElementById("f_station_assigned").value =
          u[fields.stationAssigned] || "";
        document.getElementById("f_unifi").value = u[fields.unifi] || "";
      }

      // ==================== FORM DATA COLLECTION ====================
      function collectForm() {
        // Get the raw hire date value
        const hireDateInput = document.getElementById("f_hiredate").value.trim();
        
        // Parse the date
        let hireDateValue = null;
        if (hireDateInput) {
          const parsedDate = parseDateInput(hireDateInput);
          if (parsedDate) {
            hireDateValue = formatDateForSharePoint(parsedDate);
          }
        }
        
        return {
          [fields.fullName]: document.getElementById("f_fullName").value.trim(),
          [fields.username]: document.getElementById("f_username").value.trim(),
          [fields.firstName]: document.getElementById("f_firstName").value.trim(),
          [fields.lastName]: document.getElementById("f_lastName").value.trim(),
          [fields.windowsPassword]: document.getElementById("f_winPwd").value.trim(),
          [fields.iloan]: document.getElementById("f_iloan").value.trim(),
          [fields.five9]: document.getElementById("f_five9").value.trim(),
          [fields.five9Password]: document.getElementById("f_five9Pwd").value.trim(),
          [fields.five9StationId]: document.getElementById("f_five9Station").value.trim(),
          [fields.phone]: formatPhoneNumber(document.getElementById("f_phone").value.trim()),
          [fields.department]: document.getElementById("f_department").value.trim(),
          [fields.position]: document.getElementById("f_position").value.trim(),
          [fields.location]: document.getElementById("f_location").value.trim(),
          [fields.manager]: document.getElementById("f_manager").value.trim(),
          [fields.fusionId]: document.getElementById("f_fusionId").value.trim(),
          [fields.hireDate]: hireDateValue,
          [fields.statusUser]: document.getElementById("f_status_user").value,
          [fields.statusEmail]: document.getElementById("f_status_email").value,
          [fields.statusManager]: document.getElementById("f_status_manager").value,
          [fields.stationAssigned]: document.getElementById("f_station_assigned").value,
          [fields.unifi]: document.getElementById("f_unifi").value,
        };
      }

      // ==================== SHAREPOINT API FUNCTIONS ====================
      async function createItem(payload) {
        try {
          // Validate date format if present
          if (payload[fields.hireDate] && typeof payload[fields.hireDate] === 'string') {
            // Ensure date format is correct for SharePoint
            if (!payload[fields.hireDate].endsWith('Z') && !payload[fields.hireDate].includes('T')) {
              console.warn('Date might be in wrong format, attempting to fix:', payload[fields.hireDate]);
              const parsed = parseDateInput(payload[fields.hireDate]);
              if (parsed) {
                payload[fields.hireDate] = formatDateForSharePoint(parsed);
              }
            }
          }
          
          const resp = await fetch(
            `${siteUrl}/_api/web/lists(guid'${listGuid}')/items`,
            {
              method: "POST",
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "X-RequestDigest": requestDigest,
              },
              body: JSON.stringify(payload),
            }
          );
          
          if (!resp.ok) {
            const errorText = await resp.text();
            console.error(
              "SharePoint Create Error:",
              resp.status,
              resp.statusText,
              errorText
            );
            
            // Try to parse error for better messages
            let errorMsg = "Failed to create item: " + resp.status;
            try {
              const errorJson = JSON.parse(errorText);
              if (errorJson.odata && errorJson.odata.error && errorJson.odata.error.message) {
                errorMsg += " - " + errorJson.odata.error.message.value;
              }
            } catch (e) {
              // If can't parse JSON, use text
              errorMsg += " - " + errorText.substring(0, 100);
            }
            
            throw new Error(errorMsg);
          }
          return await resp.json();
        } catch (e) {
          console.error("Create item error:", e);
          throw e;
        }
      }

      async function updateItem(id, payload) {
        try {
          // Validate date format if present
          if (payload[fields.hireDate] && typeof payload[fields.hireDate] === 'string') {
            // Ensure date format is correct for SharePoint
            if (!payload[fields.hireDate].endsWith('Z') && !payload[fields.hireDate].includes('T')) {
              console.warn('Date might be in wrong format, attempting to fix:', payload[fields.hireDate]);
              const parsed = parseDateInput(payload[fields.hireDate]);
              if (parsed) {
                payload[fields.hireDate] = formatDateForSharePoint(parsed);
              }
            }
          }
          
          const resp = await fetch(
            `${siteUrl}/_api/web/lists(guid'${listGuid}')/items(${id})`,
            {
              method: "POST",
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
                "X-RequestDigest": requestDigest,
              },
              body: JSON.stringify(payload),
            }
          );
          if (!resp.ok) {
            const errorText = await resp.text();
            console.error(
              `SharePoint Update Error (ID: ${id}):`,
              resp.status,
              resp.statusText,
              errorText
            );
            
            let errorMsg = "Failed to update item: " + resp.status;
            try {
              const errorJson = JSON.parse(errorText);
              if (errorJson.odata && errorJson.odata.error && errorJson.odata.error.message) {
                errorMsg += " - " + errorJson.odata.error.message.value;
              }
            } catch (e) {
              errorMsg += " - " + errorText.substring(0, 100);
            }
            
            throw new Error(errorMsg);
          }
        } catch (e) {
          console.error("Update item error:", e);
          throw e;
        }
      }

      // ==================== SINGLE ENTRY SUBMISSION ====================
      async function submitNewHire() {
        const payload = collectForm();
        if (!requestDigest)
          return showStatus(
            "Request digest token is missing. Please refresh the page.",
            "error"
          );
        if (
          !payload[fields.fullName] ||
          !payload[fields.username] ||
          !payload[fields.iloan]
        ) {
          return showStatus(
            "Critical fields (Full Name, M365 Username, ILOAN ID) cannot be empty.",
            "error"
          );
        }
        
        // Validate date
        if (payload[fields.hireDate] && typeof payload[fields.hireDate] === 'string') {
          const parsed = parseDateInput(payload[fields.hireDate]);
          if (!parsed) {
            return showStatus(
              "Invalid hire date format. Please use MM/DD/YYYY or YYYY-MM-DD.",
              "error"
            );
          }
        }
        
        try {
          const createdItem = await createItem(payload);
          loadedItemId = null;
          updateSingleTabButtons();
          showStatus(
            `New hire created successfully. SharePoint ID: ${createdItem.Id}`,
            "success"
          );
          if (document.getElementById("tab-record").style.display !== 'none') {
            loadAllRecords();
          }
        } catch (e) {
          console.error(e);
          showStatus(
            `Error creating new hire: ${e.message}`,
            "error"
          );
        }
      }

      async function updateHire() {
        if (!loadedItemId)
          return showStatus(
            "Load a user before attempting to update.",
            "error"
          );
        if (!requestDigest)
          await refreshDigest();
        const payload = collectForm();
        if (
          !payload[fields.fullName] ||
          !payload[fields.username] ||
          !payload[fields.iloan]
        ) {
          return showStatus(
            "Critical fields (Full Name, M365 Username, ILOAN ID) cannot be empty.",
            "error"
          );
        }

        // Validate date
        if (payload[fields.hireDate] && typeof payload[fields.hireDate] === 'string') {
          const parsed = parseDateInput(payload[fields.hireDate]);
          if (!parsed) {
            return showStatus(
              "Invalid hire date format. Please use MM/DD/YYYY or YYYY-MM-DD.",
              "error"
            );
          }
        }

        try {
          await updateItem(loadedItemId, payload);
          showStatus(
            `User (ID: ${loadedItemId}) updated successfully.`,
            "success"
          );
          // Refresh records if in record view
          if (document.getElementById("tab-record").style.display !== 'none') {
            loadAllRecords();
          }
        } catch (e) {
          console.error(e);
          showStatus(
            `Error updating user (ID: ${loadedItemId}): ${e.message}`,
            "error"
          );
        }
      }

      // ==================== SEARCH AUTOCOMPLETE ====================
      function initSearchAutocomplete() {
        const input = document.getElementById("nh-search");
        const suggestionsEl = document.getElementById("nh-search-suggestions");

        if (!input || !suggestionsEl) return;

        input.addEventListener("input", () => {
          const value = input.value.trim();
          activeSuggestionIndex = -1;

          if (searchDebounceTimer) {
            clearTimeout(searchDebounceTimer);
          }

          if (value.length < 2) {
            hideSuggestions();
            return;
          }

          searchDebounceTimer = setTimeout(() => {
            performSearchSuggestions(value);
          }, 250);
        });

        input.addEventListener("keydown", (e) => {
          if (!searchSuggestions.length) return;

          if (e.key === "ArrowDown") {
            e.preventDefault();
            activeSuggestionIndex =
              (activeSuggestionIndex + 1) % searchSuggestions.length;
            renderSuggestions();
          } else if (e.key === "ArrowUp") {
            e.preventDefault();
            activeSuggestionIndex =
              (activeSuggestionIndex - 1 + searchSuggestions.length) %
              searchSuggestions.length;
            renderSuggestions();
          } else if (e.key === "Enter") {
            if (activeSuggestionIndex >= 0) {
              e.preventDefault();
              const selected = searchSuggestions[activeSuggestionIndex];
              if (selected) {
                selectSuggestion(selected);
              }
            }
          } else if (e.key === "Escape") {
            hideSuggestions();
          }
        });

        document.addEventListener("click", (e) => {
          if (
            !suggestionsEl.contains(e.target) &&
            e.target !== input &&
            !e.target.classList.contains("search-clear-btn")
          ) {
            hideSuggestions();
          }
        });
      }

      function hideSuggestions() {
        const suggestionsEl = document.getElementById("nh-search-suggestions");
        if (suggestionsEl) {
          suggestionsEl.style.display = "none";
          suggestionsEl.innerHTML = "";
        }
        searchSuggestions = [];
        activeSuggestionIndex = -1;
      }

      async function performSearchSuggestions(query) {
        const suggestionsEl = document.getElementById("nh-search-suggestions");
        if (!suggestionsEl) return;

        const safeQuery = query.replace(/'/g, "''");

        const filter = `$filter=substringof('${safeQuery}', ${fields.username}) or substringof('${safeQuery}', ${fields.fullName}) or substringof('${safeQuery}', ${fields.iloan}) or substringof('${safeQuery}', ${fields.five9}) or substringof('${safeQuery}', ${fields.fusionId})`;

        const selectFields = ["Id"].concat(Object.values(fields));
        const endpoint = `${siteUrl}/_api/web/lists(guid'${listGuid}')/items?${filter}&$top=10&$select=${selectFields.join(
          ","
        )}`;

        try {
          const resp = await fetch(endpoint, {
            headers: { Accept: "application/json;odata=nometadata" },
          });
          if (!resp.ok) {
            hideSuggestions();
            return;
          }

          const data = await resp.json();
          if (!data.value || data.value.length === 0) {
            hideSuggestions();
            return;
          }

          searchSuggestions = data.value;
          activeSuggestionIndex = -1;
          renderSuggestions();
        } catch (e) {
          console.error("Autocomplete search error:", e);
          hideSuggestions();
        }
      }

      function renderSuggestions() {
        const suggestionsEl = document.getElementById("nh-search-suggestions");
        if (!suggestionsEl) return;

        if (!searchSuggestions.length) {
          hideSuggestions();
          return;
        }

        suggestionsEl.innerHTML = searchSuggestions
          .map((item, index) => {
            const isActive = index === activeSuggestionIndex;
            const name = item[fields.fullName] || "(No name)";
            const username = item[fields.username] || "";
            const iloan = item[fields.iloan] || "";
            const five9 = item[fields.five9] || "";
            const fusionId = item[fields.fusionId] || "";

            const metaParts = [];
            if (username) metaParts.push(username);
            if (iloan) metaParts.push("ILOAN: " + iloan);
            if (five9) metaParts.push("Five9: " + five9);
            if (fusionId) metaParts.push("FusionID: " + fusionId);

            return `
              <div class="search-suggestion-item ${
                isActive ? "active" : ""
              }" data-index="${index}">
													<span class="search-suggestion-title">${escapeHtml(
                  name
                )}</span>
													<span class="search-suggestion-meta">${escapeHtml(
                  metaParts.join(" • ")
                )}</span>
												</div>
            `;
          })
          .join("");

        suggestionsEl.style.display = "block";
        
        // Add event listeners to suggestion items
        document.querySelectorAll('.search-suggestion-item').forEach((item, index) => {
          item.addEventListener('click', () => handleSuggestionClick(index));
        });
      }

      function handleSuggestionClick(index) {
        const item = searchSuggestions[index];
        if (item) {
          selectSuggestion(item);
        }
      }

      function selectSuggestion(item) {
        loadedItemId = item.Id;
        populateForm(item);
        updateSingleTabButtons();
        const input = document.getElementById("nh-search");
        if (input) {
          input.value = item[fields.username] || item[fields.fullName] || "";
        }
        hideSuggestions();
        showStatus(
          "User loaded from search suggestions (ID: " + loadedItemId + ").",
          "success"
        );
      }

      function clearSearch() {
        const input = document.getElementById("nh-search");
        if (input) {
          input.value = "";
          input.focus();
        }
        hideSuggestions();
        loadedItemId = null;
        updateSingleTabButtons();
      }

      // ==================== CSV PARSING ====================
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
        return result.map((v) => v.trim());
      }

      function parseCsv(text) {
        const lines = text.split(/\r?\n/).filter((l) => l.trim() !== "");
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
        const get = (key) => (normRow[key] || "").trim();

        const fullName = get("fullname");
        const firstName = get("firstname");
        const lastName = get("lastname");
        const username = get("m365usernamecomputerusername");
        const winPwd = get("windowsm365password");
        const five9User = get("five9username");
        const five9Pwd = get("five9password");
        const five9Station = get("five9stationid");
        const iloan = get("iloanid");
        const dept = get("department");
        const position = get("position");
        const hireDateRaw = get("hiredate");
        const location = get("location");
        const manager = get("manager");
        const phone = get("phonenumber");
        const fusionId = get("fusionid");

        const payload = {};

        payload[fields.fullName] = fullName;
        payload[fields.firstName] = firstName;
        payload[fields.lastName] = lastName;
        payload[fields.username] = username;
        payload[fields.windowsPassword] = winPwd;
        payload[fields.five9] = five9User;
        payload[fields.five9Password] = five9Pwd;
        payload[fields.five9StationId] = five9Station;
        payload[fields.iloan] = iloan;
        payload[fields.department] = dept;
        payload[fields.position] = position;
        payload[fields.location] = location;
        payload[fields.manager] = manager;
        payload[fields.phone] = formatPhoneNumber(phone);
        payload[fields.fusionId] = fusionId;

        // Parse hire date
        if (hireDateRaw) {
          const parsedDate = parseDateInput(hireDateRaw);
          if (parsedDate) {
            payload[fields.hireDate] = formatDateForSharePoint(parsedDate);
          } else {
            payload[fields.hireDate] = null;
          }
        } else {
          payload[fields.hireDate] = null;
        }

        return payload;
      }

      async function findExistingUserId(username, iloan, five9, fusionId) {
        let filter = "";
        function esc(v) {
          return v.replace(/'/g, "''");
        }

        if (username) {
          filter = `${fields.username} eq '${esc(username)}'`;
        } else if (iloan) {
          filter = `${fields.iloan} eq '${esc(iloan)}'`;
        } else if (five9) {
          filter = `${fields.five9} eq '${esc(five9)}'`;
        } else if (fusionId) {
          filter = `${fields.fusionId} eq '${esc(fusionId)}'`;
        } else {
          return null;
        }

        const endpoint = `${siteUrl}/_api/web/lists(guid'${listGuid}')/items?$top=1&$filter=${filter}`;

        try {
          const resp = await fetch(endpoint, {
            headers: { Accept: "application/json;odata=nometadata" },
          });
          if (!resp.ok) return null;

          const data = await resp.json();
          if (!data.value || data.value.length === 0) return null;

          return data.value[0].Id;
        } catch (e) {
          console.error("Error checking existing user", e);
          return null;
        }
      }

      function logCsv(msg) {
        const logEl = document.getElementById("csvLog");
        logEl.textContent += msg + "\n";
        logEl.scrollTop = logEl.scrollHeight;
      }

      async function handleCsvFile() {
        const input = document.getElementById("csvFile");
        if (!input.files || !input.files.length) {
          showStatus("Select a CSV file first.", "error");
          return;
        }
        showStatus("", "success");

        const file = input.files[0];
        const reader = new FileReader();
        reader.onload = async (e) => {
          document.getElementById("csvLog").textContent = "";
          logCsv("Reading CSV: " + file.name);

          const text = e.target.result;
          const parsed = parseCsv(text);
          logCsv("Parsed rows (excluding header): " + parsed.length);

          csvRows = [];

          let idx = 0;
          for (const normRow of parsed) {
            const payload = mapCsvRowToPayload(normRow);

            const username = payload[fields.username];
            const iloan = payload[fields.iloan];
            const five9 = payload[fields.five9];
            const fusionId = payload[fields.fusionId];

            const existingId = await findExistingUserId(
              username,
              iloan,
              five9,
              fusionId
            );

            csvRows.push({
              index: idx,
              payload,
              existingId,
              include: true,
              editing: false,
            });

            logCsv(
              `Row ${idx + 1}: ${
                payload[fields.fullName] || "(no name)"
              } - ` + (existingId ? `EXISTING (ID ${existingId})` : "NEW")
            );

            idx++;
          }

          renderCsvPreview();
        };
        reader.readAsText(file);
      }

      function renderCsvPreview() {
        const body = document.getElementById("csvPreviewBody");
        const summary = document.getElementById("csvSummary");

        if (!csvRows.length) {
          body.innerHTML = "";
          summary.textContent = "No rows loaded.";
          return;
        }

        let newCount = 0;
        let existingCount = 0;

        const rowsHtml = csvRows
          .map((r, i) => {
            if (r.existingId) existingCount++;
            else newCount++;

            const isCriticalMissing =
              !r.payload[fields.fullName] ||
              !r.payload[fields.username] ||
              !r.payload[fields.iloan];
            const rowClass = isCriticalMissing ? "missing-critical" : "";
            const editRowClass = r.editing ? "edit-row" : "";

            if (!r.editing) {
              return `
                <tr class="${rowClass} ${editRowClass}">
												<td>${i + 1}</td>
												<td>
													<input type="checkbox" ${r.include ? "checked" : ""} data-index="${i}" />
												</td>
												<td>
													<span style="font-weight: 600; color: ${
                      r.existingId ? "#0078d4" : "#333"
                    };">${
                r.existingId
                  ? "Existing (ID " + r.existingId + ")"
                  : "New"
              }</span>
												</td>
												<td>${escapeHtml(r.payload[fields.fullName] || "")}</td>
												<td>${escapeHtml(r.payload[fields.firstName] || "")}</td>
												<td>${escapeHtml(r.payload[fields.lastName] || "")}</td>
												<td>${escapeHtml(r.payload[fields.username] || "")}</td>
												<td>${escapeHtml(r.payload[fields.windowsPassword] || "")}</td>
												<td>${escapeHtml(r.payload[fields.five9] || "")}</td>
												<td>${escapeHtml(r.payload[fields.five9Password] || "")}</td>
												<td>${escapeHtml(r.payload[fields.five9StationId] || "")}</td>
												<td>${escapeHtml(r.payload[fields.iloan] || "")}</td>
												<td>${escapeHtml(r.payload[fields.fusionId] || "")}</td>
												<td>${escapeHtml(r.payload[fields.department] || "")}</td>
												<td>${escapeHtml(r.payload[fields.position] || "")}</td>
												<td>${escapeHtml(formatDateForDisplay(r.payload[fields.hireDate] || ""))}</td>
												<td>${escapeHtml(r.payload[fields.location] || "")}</td>
												<td>${escapeHtml(r.payload[fields.manager] || "")}</td>
												<td>${escapeHtml(r.payload[fields.phone] || "")}</td>
												<td>
													<button style="font-size: 11px; padding: 4px 8px;" class="btn-secondary edit-csv-row-btn" data-index="${i}">Edit</button>
												</td>
											</tr>
              `;
            }

            return `
              <tr class="${editRowClass}">
											<td>${i + 1}</td>
											<td>
												<input type="checkbox" ${r.include ? "checked" : ""} data-index="${i}" />
											</td>
											<td>
												<span style="font-weight: 600; color: ${
                    r.existingId ? "#0078d4" : "#333"
                  };">${
              r.existingId ? "Existing (ID " + r.existingId + ")" : "New"
            }</span>
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="fullName" value="${escapeHtml(
              r.payload[fields.fullName] || ""
            )}" ${!r.payload[fields.fullName] ? 'style="border-color: #ff4d4d;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="firstName" value="${escapeHtml(
              r.payload[fields.firstName] || ""
            )}" ${!r.payload[fields.firstName] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="lastName" value="${escapeHtml(
              r.payload[fields.lastName] || ""
            )}" ${!r.payload[fields.lastName] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="username" value="${escapeHtml(
              r.payload[fields.username] || ""
            )}" ${!r.payload[fields.username] ? 'style="border-color: #ff4d4d;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="windowsPassword" value="${escapeHtml(
              r.payload[fields.windowsPassword] || ""
            )}" ${!r.payload[fields.windowsPassword] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="five9" value="${escapeHtml(
              r.payload[fields.five9] || ""
            )}" ${!r.payload[fields.five9] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="five9Password" value="${escapeHtml(
              r.payload[fields.five9Password] || ""
            )}" ${!r.payload[fields.five9Password] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="five9StationId" value="${escapeHtml(
              r.payload[fields.five9StationId] || ""
            )}" ${!r.payload[fields.five9StationId] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="iloan" value="${escapeHtml(
              r.payload[fields.iloan] || ""
            )}" ${!r.payload[fields.iloan] ? 'style="border-color: #ff4d4d;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="fusionId" value="${escapeHtml(
              r.payload[fields.fusionId] || ""
            )}" />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="department" value="${escapeHtml(
              r.payload[fields.department] || ""
            )}" ${!r.payload[fields.department] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="position" value="${escapeHtml(
              r.payload[fields.position] || ""
            )}" ${!r.payload[fields.position] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input type="text" class="edit-field" data-index="${i}" data-field="hireDate" value="${escapeHtml(
              formatDateForDisplay(r.payload[fields.hireDate] || "")
            )}" ${!r.payload[fields.hireDate] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="location" value="${escapeHtml(
              r.payload[fields.location] || ""
            )}" ${!r.payload[fields.location] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="manager" value="${escapeHtml(
              r.payload[fields.manager] || ""
            )}" ${!r.payload[fields.manager] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<input class="edit-field" data-index="${i}" data-field="phone" value="${escapeHtml(
              r.payload[fields.phone] || ""
            )}" ${!r.payload[fields.phone] ? 'style="border-color: #ffcc00;"' : ''} />
											</td>
											<td>
												<button style="font-size: 10px; padding: 2px 4px; margin-bottom: 3px;" class="btn-primary save-csv-row-btn" data-index="${i}">Save</button>
												<button style="font-size: 10px; padding: 2px 4px;" class="btn-secondary cancel-csv-row-btn" data-index="${i}">Cancel</button>
											</td>
										</tr>
            `;
          })
          .join("");

        body.innerHTML = rowsHtml;
        summary.textContent = `Loaded ${csvRows.length} row(s). New: ${newCount}, Existing: ${existingCount}. (Highlighted rows are missing critical data.)`;
        
        // Add event listeners for checkboxes and buttons
        document.querySelectorAll('input[type="checkbox"][data-index]').forEach(checkbox => {
          const index = parseInt(checkbox.getAttribute('data-index'));
          checkbox.addEventListener('change', () => toggleCsvRowInclude(index, checkbox.checked));
        });
        
        document.querySelectorAll('.edit-csv-row-btn').forEach(btn => {
          const index = parseInt(btn.getAttribute('data-index'));
          btn.addEventListener('click', () => editCsvRow(index));
        });
        
        document.querySelectorAll('.save-csv-row-btn').forEach(btn => {
          const index = parseInt(btn.getAttribute('data-index'));
          btn.addEventListener('click', () => saveCsvRow(index));
        });
        
        document.querySelectorAll('.cancel-csv-row-btn').forEach(btn => {
          const index = parseInt(btn.getAttribute('data-index'));
          btn.addEventListener('click', () => cancelCsvRow(index));
        });
        
        document.querySelectorAll('.edit-field').forEach(input => {
          input.addEventListener('input', () => {
            const index = parseInt(input.getAttribute('data-index'));
            const field = input.getAttribute('data-field');
            if (field === 'hireDate') {
              const parsed = parseDateInput(input.value);
              if (parsed) {
                csvRows[index].payload[fields[field]] = formatDateForSharePoint(parsed);
              } else {
                csvRows[index].payload[fields[field]] = null;
              }
            } else if (field === 'phone') {
              csvRows[index].payload[fields[field]] = formatPhoneNumber(input.value.trim());
            } else {
              csvRows[index].payload[fields[field]] = input.value.trim();
            }
          });
        });
      }

      function toggleCsvRowInclude(index, checked) {
        if (csvRows[index]) {
          csvRows[index].include = checked;
        }
      }

      function editCsvRow(index) {
        csvRows.forEach((r, i) => (r.editing = i === index));
        renderCsvPreview();
      }

      function cancelCsvRow(index) {
        csvRows[index].editing = false;
        renderCsvPreview();
      }

      function saveCsvRow(index) {
        csvRows[index].editing = false;
        renderCsvPreview();
        logCsv(`Row ${index + 1} updated inline.`);
      }

      // ==================== CSV IMPORT ====================
      async function importCsvRows() {
        const selectedRows = csvRows.filter((r) => r.include);
        if (selectedRows.length === 0) {
          showStatus("No rows selected for import.", "error");
          return;
        }
        if (!requestDigest) {
          showStatus(
            "Request digest token is missing. Please refresh the page before import.",
            "error"
          );
          return;
        }

        logCsv("Starting import of " + selectedRows.length + " selected rows...");

        let success = 0;
        let fail = 0;

        for (const row of selectedRows) {
          const name = row.payload[fields.fullName] || "(no name)";
          const i = row.index;

          if (
            !row.payload[fields.fullName] ||
            !row.payload[fields.username] ||
            !row.payload[fields.iloan]
          ) {
            logCsv(`Row ${i + 1}: SKIPPED (Missing critical data).`);
            fail++;
            continue;
          }

          try {
            if (row.existingId) {
              await updateItem(row.existingId, row.payload);
              logCsv(
                `Row ${i + 1}: Updated existing user "${name}" (ID ${
                  row.existingId
                }).`
              );
            } else {
              const created = await createItem(row.payload);
              logCsv(
                `Row ${i + 1}: Created new user "${name}" (ID ${created.Id}).`
              );
            }
            success++;
          } catch (e) {
            console.error(e);
            logCsv(`Row ${i + 1}: ERROR for "${name}" - ${e.message}`);
            fail++;
          }
        }

        logCsv(`Import complete. Success: ${success}, Failed: ${fail}.`);
        showStatus(
          `CSV import complete. Success: ${success}, Failed: ${fail}.`,
          fail ? "error" : "success"
        );
        
        // Refresh records if in record view
        if (document.getElementById("tab-record").style.display !== 'none') {
          loadAllRecords();
        }
      }

      // ==================== STATUS MESSAGES ====================
      function showStatus(msg, type) {
        const box = document.getElementById("statusBox");
        if (!msg) {
          box.style.display = "none";
          box.textContent = "";
          return;
        }

        box.textContent = msg;
        box.className = "status-box " + type;
        box.style.display = "block";
      }

      // ==================== RECORD VIEW ACTION FUNCTIONS ====================

      async function markSelectedField(fieldName, value, label) {
        if (selectedIds.size === 0) return;
        if (!requestDigest) await refreshDigest();

        let updated = 0;
        let failed = 0;

        for (const id of selectedIds) {
          try {
            const resp = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listGuid}')/items(${id})`,
              {
                method: "POST",
                headers: {
                  Accept: "application/json;odata=nometadata",
                  "Content-Type": "application/json;odata=nometadata",
                  "X-HTTP-Method": "MERGE",
                  "IF-MATCH": "*",
                  "X-RequestDigest": requestDigest,
                },
                body: JSON.stringify({ [fieldName]: value }),
              }
            );
            if (resp.ok) {
              const rec = allRecords.find(r => r.Id === id);
              if (rec) rec[fieldName] = value;
              updated++;
            } else {
              failed++;
            }
          } catch (e) {
            console.error(`Error marking ${label} for ID ${id}:`, e);
            failed++;
          }
        }

        const msg = failed === 0
          ? `${label}: marked ${updated} record(s) successfully.`
          : `${label}: ${updated} updated, ${failed} failed.`;
        showStatus(msg, failed === 0 ? 'success' : 'error');

        selectedIds.clear();
        updateActionBar();
        updateItStats();
        renderRecordsTable();
        updateRecordCount();
      }

      function markM365Profile()   { markSelectedField(fields.profileCreated,  'YES', 'M365 Profile Created'); }
      function markUnifiProfile()  { markSelectedField(fields.unifi,            'Yes', 'Unifi Profile Created'); }
      function markAssignManager() { markSelectedField(fields.statusManager,    'YES', 'Manager Assigned'); }
      function markSendEmail()     { markSelectedField(fields.statusEmail,      'YES', 'Welcome Email Sent'); }
      function markAssignStation() { markSelectedField(fields.stationAssigned,  'Yes', 'Station Assigned'); }

      // ==================== SINGLE ENTRY BUTTON STATE ====================

      function updateSingleTabButtons() {
        const saveBtn   = document.getElementById('saveNewHireBtn');
        const updateBtn = document.getElementById('updateHireBtn');
        if (!saveBtn || !updateBtn) return;
        if (loadedItemId) {
          saveBtn.style.display   = 'none';
          updateBtn.style.display = '';
        } else {
          saveBtn.style.display   = '';
          updateBtn.style.display = 'none';
        }
      }

      function showSingleConfirm(message, fn) {
        singleConfirmPendingFn = fn;
        const msg = document.getElementById('singleConfirmMsg');
        if (msg) msg.textContent = message;
        document.getElementById('singleConfirmOverlay').classList.add('visible');
        document.getElementById('singleConfirmModal').classList.add('visible');
      }

      function hideSingleConfirm() {
        singleConfirmPendingFn = null;
        document.getElementById('singleConfirmOverlay').classList.remove('visible');
        document.getElementById('singleConfirmModal').classList.remove('visible');
      }

      function executeSingleConfirm() {
        const fn = singleConfirmPendingFn;
        hideSingleConfirm();
        if (fn) fn();
      }

      // ==================== DELETE SELECTED ====================

      let pendingDeleteIds = [];

      function showDeleteModal() {
        if (selectedIds.size === 0) return;
        pendingDeleteIds = Array.from(selectedIds);
        const msg = document.getElementById('deleteModalMsg');
        if (msg) {
          msg.textContent = `Are you sure you want to permanently delete ${pendingDeleteIds.length} record(s)? This cannot be undone.`;
        }
        document.getElementById('deleteModalOverlay').classList.add('visible');
        document.getElementById('deleteModal').classList.add('visible');
      }

      function hideDeleteModal() {
        pendingDeleteIds = [];
        document.getElementById('deleteModalOverlay').classList.remove('visible');
        document.getElementById('deleteModal').classList.remove('visible');
      }

      async function performDelete() {
        if (pendingDeleteIds.length === 0) { hideDeleteModal(); return; }
        const idsToDelete = [...pendingDeleteIds]; // snapshot before hideDeleteModal clears the array
        hideDeleteModal();
        await refreshDigest();

        let deleted = 0;
        let failed  = 0;

        for (const id of idsToDelete) {
          try {
            const resp = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listGuid}')/items(${id})`,
              {
                method: "POST",
                headers: {
                  Accept: "application/json;odata=nometadata",
                  "Content-Type": "application/json;odata=nometadata",
                  "X-HTTP-Method": "DELETE",
                  "IF-MATCH": "*",
                  "X-RequestDigest": requestDigest,
                },
              }
            );
            if (resp.ok) {
              allRecords = allRecords.filter(r => r.Id !== id);
              deleted++;
            } else {
              failed++;
            }
          } catch (e) {
            console.error(`Error deleting ID ${id}:`, e);
            failed++;
          }
        }

        selectedIds.clear();
        updateActionBar();
        extractUniqueValues();
        updateItStats();
        applyFilters();

        showStatus(
          failed === 0
            ? `Deleted ${deleted} record(s) successfully.`
            : `Deleted ${deleted}, failed ${failed}.`,
          failed === 0 ? 'success' : 'error'
        );
      }
