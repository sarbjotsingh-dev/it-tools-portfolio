       // ====== RANDOM QUOTES LIBRARY ======
        const recognitionQuotes = [
            {
                text: "Recognition is the greatest motivator. A little appreciation can change a day, even change a life.",
                author: "Unknown"
            },
            {
                text: "People work for money but go the extra mile for recognition, praise and rewards.",
                author: "Dale Carnegie"
            },
            {
                text: "Appreciation is a wonderful thing: It makes what is excellent in others belong to us as well.",
                author: "Voltaire"
            },
            {
                text: "A sincere thank you is the simplest form of recognition and one of the most powerful.",
                author: "Unknown"
            },
            {
                text: "When people are financially invested, they want a return. When people are emotionally invested, they want to contribute.",
                author: "Simon Sinek"
            },
            {
                text: "Recognition is not a scarce resource. You can't use it up or run out of it.",
                author: "Susan M. Heathfield"
            },
            {
                text: "The deepest principle in human nature is the craving to be appreciated.",
                author: "William James"
            },
            {
                text: "Make it a rule to thank anyone who helps you, regardless of their role or the size of their contribution.",
                author: "Adam Grant"
            },
            {
                text: "Great things in business are never done by one person. They're done by a team of people.",
                author: "Steve Jobs"
            },
            {
                text: "Recognition is the fuel that drives innovation, engagement, and loyalty.",
                author: "Unknown"
            },
            {
                text: "Feeling valued and appreciated is a fundamental human need. It's what motivates us to do our best work.",
                author: "Daniel Pink"
            },
            {
                text: "The best way to appreciate someone's contribution is to acknowledge it publicly and specifically.",
                author: "Ken Blanchard"
            },
            {
                text: "A simple 'thank you' can have a ripple effect that extends far beyond the moment.",
                author: "Unknown"
            },
            {
                text: "Recognition isn't about grand gestures; it's about noticing and appreciating the everyday contributions.",
                author: "Unknown"
            },
            {
                text: "The currency of real networking is not greed but generosity.",
                author: "Keith Ferrazzi"
            }
        ];

        // Function to display random quote
        function showRandomQuote() {
            const quoteContainer = document.getElementById('random-quote');
            if (quoteContainer) {
                const randomIndex = Math.floor(Math.random() * recognitionQuotes.length);
                const randomQuote = recognitionQuotes[randomIndex];
                
                quoteContainer.innerHTML = `"${randomQuote.text}" 
                    <span class="quote-author">- ${randomQuote.author}</span>`;
            }
        }

        // ====== SHAREPOINT CONFIGURATION ======
        const listGuid = "13e0a045-43fe-470e-85ba-36d59027b341";
        const baseUrl = "https://your-tenant.sharepoint.com";
        const listItemType = "SP.Data.Virtual_x0020_High_x0020_FiveListItem";

        // ====== GLOBAL STATE ======
        let currentUserId = null;
        let currentUserName = null;
        let dataMode = 'count';
        let allItems = [];
        let chartDept, chartLoc, chartMonth;

        // ====== GLOBAL FUNCTIONS ======
        
        // --- FILL-IN VISIBILITY LOGIC ---
        function checkFill(select, fieldInternalName) {
            const fillBox = document.getElementById(`fill-${fieldInternalName}`);
            const textInput = document.getElementById(`txt-${fieldInternalName}`);
            if (select.value === "FILL_IN_VAL") {
                fillBox.style.display = 'block';
                textInput.required = true;
            } else {
                fillBox.style.display = 'none';
                textInput.required = false;
            }
        }

        // --- FORM CHOICES SYNC ---
        async function syncFormChoices() {
            const fields = ["field_2", "field_8", "field_9", "field_5"];
            const fillInEnabled = ["field_2", "field_8", "field_5"];

            for (let fName of fields) {
                const url = `${baseUrl}/_api/web/lists(guid'${listGuid}')/fields?$filter=InternalName eq '${fName}'`;
                try {
                    const resp = await fetch(url, { headers: { "Accept": "application/json;odata=nometadata" } });
                    const data = await resp.json();
                    if (data.value && data.value[0].Choices) {
                        let html = data.value[0].Choices.map(c => `<option value="${c}">${c}</option>`).join('');
                        
                        if (fillInEnabled.includes(fName)) {
                            html += `<option value="FILL_IN_VAL">-- Other (Specify) --</option>`;
                        }
                        
                        document.getElementById(`sel-${fName}`).innerHTML = html;
                    }
                } catch (e) { console.error("Sync error", e); }
            }
        }

        // --- PEOPLE PICKER FUNCTIONS ---
        async function getDigest() {
            const resp = await fetch(`${baseUrl}/_api/contextinfo`, { 
                method: "POST", 
                headers: { "Accept": "application/json;odata=nometadata" } 
            });
            const data = await resp.json();
            return data.FormDigestValue;
        }

        async function pickUser(name, key) {
            const empInput = document.getElementById('form-employee');
            const empIdInput = document.getElementById('form-employee-id');
            const suggestBox = document.getElementById('employee-suggestions');
            
            empInput.value = name;
            suggestBox.style.display = 'none';
            
            try {
                const digest = await getDigest();
                const resp = await fetch(`${baseUrl}/_api/web/ensureuser('${encodeURIComponent(key)}')`, {
                    method: 'POST', 
                    headers: { 
                        "Accept": "application/json;odata=nometadata", 
                        "X-RequestDigest": digest 
                    }
                });
                
                if (resp.ok) {
                    const data = await resp.json();
                    empIdInput.value = data.Id;
                    console.log("Selected user ID:", data.Id);
                } else {
                    console.error("Failed to ensure user");
                }
            } catch (error) {
                console.error("Error ensuring user:", error);
            }
        }

        // Initialize people picker event listener
        function initPeoplePicker() {
            const empInput = document.getElementById('form-employee');
            const suggestBox = document.getElementById('employee-suggestions');
            
            if (!empInput || !suggestBox) return;
            
            empInput.addEventListener('input', async function() {
                const val = this.value;
                if (val.length < 3) { 
                    suggestBox.style.display = 'none'; 
                    return; 
                }

                try {
                    const digest = await getDigest();
                    const searchUrl = `${baseUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
                    const resp = await fetch(searchUrl, {
                        method: "POST",
                        body: JSON.stringify({ 
                            "queryParams": { 
                                "QueryString": val, 
                                "MaximumEntitySuggestions": 8, 
                                "PrincipalType": 1 
                            } 
                        }),
                        headers: { 
                            "Accept": "application/json;odata=nometadata", 
                            "Content-Type": "application/json;odata=verbose", 
                            "X-RequestDigest": digest 
                        }
                    });
                    
                    const result = await resp.json();
                    const results = JSON.parse(result.value);
                    
                    if (results && results.length > 0) {
                        suggestBox.innerHTML = results.map(u => {
                            const displayName = u.DisplayText.replace(/'/g, "\\'");
                            const key = u.Key.replace(/'/g, "\\'");
                            return `<div class="suggestion-item" onclick="pickUser('${displayName}', '${key}')">
                                ${u.DisplayText}<br>
										<small>${u.Description || ''}</small>
									</div>`;
                        }).join('');
                        suggestBox.style.display = 'block';
                    } else {
                        suggestBox.innerHTML = '<div class="suggestion-item" style="color: #666; cursor: default;">No users found</div>';
                        suggestBox.style.display = 'block';
                    }
                } catch (error) {
                    console.error("People picker error:", error);
                    suggestBox.innerHTML = '<div class="suggestion-item" style="color: #c00; cursor: default;">Search error</div>';
                    suggestBox.style.display = 'block';
                }
            });
            
            document.addEventListener('click', function(event) {
                if (!empInput.contains(event.target) && !suggestBox.contains(event.target)) {
                    suggestBox.style.display = 'none';
                }
            });
        }

        // --- GET CURRENT USER ---
        async function getCurrentUser() {
            try {
                const response = await fetch(`${baseUrl}/_api/web/currentuser`, {
                    headers: { "Accept": "application/json;odata=nometadata" }
                });
                
                if (response.ok) {
                    const userData = await response.json();
                    currentUserId = userData.Id;
                    currentUserName = userData.Title;
                    
                    // Show initial status with user info
                    showStatus("Loading data for " + currentUserName + "...", false);
                    
                    console.log("Current user:", currentUserName, "ID:", currentUserId);
                    return userData;
                }
            } catch (error) {
                console.error("Error getting current user:", error);
                showStatus("Error getting user information", true);
            }
            return null;
        }

        // --- FORM SUBMISSION ---
        function initFormSubmission() {
            const form = document.getElementById('highFiveForm');
            if (!form) return;
            
            // Add form validation for reason length
            const reasonInput = document.getElementById('form-reason');
            if (reasonInput) {
                reasonInput.addEventListener('input', function() {
                    if (this.value.length < 20) {
                        this.setCustomValidity('Please provide at least 20 characters describing the achievement.');
                    } else {
                        this.setCustomValidity('');
                    }
                });
            }
            
            form.onsubmit = async (e) => {
                e.preventDefault();
                
                // Validate reason length
                if (reasonInput && reasonInput.value.length < 20) {
                    showStatus("Please provide a more detailed reason (at least 20 characters).", true);
                    reasonInput.focus();
                    return;
                }
                
                const employeeId = document.getElementById('form-employee-id').value;
                if(!employeeId) { 
                    showStatus("Please search and select an employee.", true);
                    return; 
                }

                const getVal = (id) => {
                    const sel = document.getElementById(`sel-${id}`).value;
                    return (sel === "FILL_IN_VAL") ? document.getElementById(`txt-${id}`).value : sel;
                };

                try {
                    const digest = await getDigest();
                    const payload = {
                        "__metadata": { "type": listItemType },
                        "field_1": document.getElementById('form-date').value,
                        "field_2": getVal('field_2'),
                        "field_5": getVal('field_5'),
                        "field_8": getVal('field_8'),
                        "field_9": document.getElementById('sel-field_9').value,
                        "field_10": document.getElementById('form-reason').value,
                        "NameofEmployeeId": parseInt(employeeId)
                    };

                    console.log("Submitting:", payload);
                    showStatus("Submitting High Five...", false);

                    const response = await fetch(`${baseUrl}/_api/web/lists(guid'${listGuid}')/items`, {
                        method: "POST",
                        body: JSON.stringify(payload),
                        headers: { 
                            "Accept": "application/json;odata=nometadata", 
                            "Content-Type": "application/json;odata=verbose", 
                            "X-RequestDigest": digest 
                        }
                    });

                    if (response.ok) {
                        showStatus("✓ High Five Submitted Successfully! Refreshing dashboard...", false);
                        
                        // Reset form but keep date as today
                        form.reset();
                        document.getElementById('form-date').valueAsDate = new Date();
                        document.getElementById('form-employee-id').value = '';
                        
                        // Show new random quote after submission
                        showRandomQuote();
                        
                        await loadData();
                        
                        setTimeout(() => showStatus(""), 3000);
                    } else {
                        const error = await response.json();
                        console.error("Submission error:", error);
                        showStatus("Error submitting. Please try again.", true);
                    }
                } catch (error) {
                    console.error("Network error:", error);
                    showStatus("Network error. Please check connection.", true);
                }
            };
        }

        // ====== DASHBOARD FUNCTIONS ======

        // Utility Functions
        function showStatus(msg, isError = false) {
            const statusEl = document.getElementById("vhf-status");
            if (!statusEl) return;
            
            if (msg) {
                // Show the combined section
                statusEl.style.display = 'flex';
                
                // Format the message with current user name
                let statusText = msg;
                if (currentUserName && !msg.includes("Loading") && !isError) {
                    if (msg.includes("Data loaded successfully")) {
                        statusText = `✓ Data loaded successfully for ${currentUserName}`;
                    } else if (msg.includes("High Five Submitted")) {
                        statusText = msg;
                    }
                }
                
                statusEl.innerHTML = `
                    <div class="status-message ${isError ? 'status-error' : 'status-success'}">
                        ${statusText}
                        ${currentUserName && !isError ? '<span class="current-user-badge">Your Data</span>' : ''}
                    </div>
                `;
            } else {
                statusEl.style.display = 'none';
                statusEl.innerHTML = "";
            }
        }

        function unique(arr) {
            return Array.from(new Set(arr)).filter(Boolean).sort();
        }
        
        function filterCurrentMonth(items) {
            const now = new Date();
            const currentMonth = now.getMonth();
            const currentYear = now.getFullYear();
            return items.filter(function(i) {
                return i.date.getMonth() === currentMonth && i.date.getFullYear() === currentYear;
            });
        }
        
        function filterByMonthRange(items, range) {
            const now = new Date();
            let startDate = new Date();
            
            switch(range) {
                case 'current':
                    startDate = new Date(now.getFullYear(), now.getMonth(), 1);
                    break;
                case 'last':
                    startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                    break;
                case 'last3':
                    startDate = new Date(now.getFullYear(), now.getMonth() - 3, 1);
                    break;
                case 'last6':
                    startDate = new Date(now.getFullYear(), now.getMonth() - 6, 1);
                    break;
                case 'last12':
                    startDate = new Date(now.getFullYear(), now.getMonth() - 12, 1);
                    break;
                default:
                    return items; // All time
            }
            
            return items.filter(function(i) {
                return i.date >= startDate;
            });
        }

        function normalizeDept(raw) {
            if (!raw) return "OTHER";
            let d = String(raw).trim().toUpperCase();

            const map = {
                "COMPLIANCE": "COMPLIANCE",
                "COMPLINACE": "COMPLIANCE",
                "COMPLAINCE": "COMPLIANCE",
                "COMPILANCE": "COMPLIANCE",
                "CSR": "CSR",
                "CUSTOMER SERVICE REPRESENTATIVE": "CSR",
                "ASSESSMENT": "ASSESSMENT",
                "ASSESMENT": "ASSESSMENT",
                "ASSESSMENTS": "ASSESSMENT",
                "QUALITY ASSURANCE": "QUALITY ASSURANCE",
                "QUALITY ASSUSRANCE": "QUALITY ASSURANCE",
                "QUALITY ASSURENCE": "QUALITY ASSURANCE",
                "QUALITY ASSURRANCE": "QUALITY ASSURANCE",
                "QA": "QUALITY ASSURANCE",
                "Q A": "QUALITY ASSURANCE",
                "COLLECTIONS": "COLLECTIONS",
                "COLLECTION": "COLLECTIONS",
                "OPERATIONS": "OPERATIONS",
                "OPERATION": "OPERATIONS",
                "IT": "IT",
                "I T": "IT",
                "FRONT DESK": "FRONT DESK",
                "FRONTDESK": "FRONT DESK",
                "MANAGEMENT": "MANAGEMENT",
                "MGMT": "MANAGEMENT"
            };

            return map[d] || d || "OTHER";
        }

        // Data Loading
        async function loadData() {
            showStatus("Loading data from SharePoint list...", false);

            const endpoint = baseUrl + "/_api/web/lists(guid'" + listGuid + "')/items"
                + "?$top=5000&$orderby=Created desc"
                + "&$select=Id,field_1,field_2,field_5,field_8,field_9,field_10,Created,NameofEmployee/Title,Author/Title&$expand=NameofEmployee,Author";

            try {
                const resp = await fetch(endpoint, {
                    headers: { Accept: "application/json;odata=nometadata" }
                });

                if (!resp.ok) {
                    showStatus("Error loading data: HTTP " + resp.status + ". Check list GUID and SharePoint permissions.", true);
                    return;
                }

                const data = await resp.json();

                allItems = data.value.map(function(i) {
                    var deptClean = normalizeDept(i.field_8 || "");
                    var highfives = parseInt(i.field_5) || 0;
                    var dateValue = i.field_1 ? new Date(i.field_1) : new Date(i.Created);
                    
                    return {
                        highfives: highfives,
                        team: i.field_2 || "N/A",
                        dept: deptClean,
                        loc: i.field_9 || "N/A",
                        reason: i.field_10 || "",
                        employee: i.NameofEmployee && i.NameofEmployee.Title ? i.NameofEmployee.Title : "Unknown Employee",
                        supervisor: i.Author && i.Author.Title ? i.Author.Title : "Unknown Supervisor",
                        date: dateValue
                    };
                });

                document.getElementById('vhf-last-refresh').textContent = new Date().toLocaleString();
                showStatus("Data loaded successfully.", false);
                initFilters();
                renderAll();

            } catch (error) {
                console.error("Fetch error:", error);
                showStatus("Network or Fetch Error: The API call failed. Ensure you are running this code within the same SharePoint domain.", true);
            }
        }

        // Filter Initialization
        function initFilters() {
            function fill(select, arr) {
                // Clear existing options except the first one
                while (select.options.length > 1) {
                    select.remove(1);
                }
                
                unique(arr).forEach(function(v) {
                    var opt = document.createElement("option");
                    opt.value = v;
                    opt.textContent = v;
                    select.appendChild(opt);
                });
            }

            fill(document.getElementById("vhf-filter-dept"), allItems.map(function(i) { return i.dept; }));
            fill(document.getElementById("vhf-filter-location"), allItems.map(function(i) { return i.loc; }));

            var datalist = document.getElementById("vhf-search-list");
            datalist.innerHTML = "";
            unique(allItems.map(function(i) { return i.employee; })).forEach(function(name) {
                datalist.innerHTML += '<option value="' + name + '">';
            });
        }

        // Filter Application - UPDATED with Month filter fix
        function applyFilters() {
            var items = allItems.slice();

            // FILTER 1: Auto-filter to current user's data (as supervisor)
            if (currentUserName) {
                items = items.filter(function(i) { 
                    return i.supervisor === currentUserName; 
                });
            }

            // FILTER 2: Search employee (recipient)
            const filterSearch = document.getElementById("vhf-filter-search");
            if (filterSearch.value) {
                var s = filterSearch.value.toLowerCase();
                items = items.filter(function(i) {
                    return (i.employee && i.employee.toLowerCase().includes(s));
                });
            }

            // FILTER 3: Department
            const filterDept = document.getElementById("vhf-filter-dept");
            if (filterDept.value) {
                items = items.filter(function(i) { return i.dept === filterDept.value; });
            }

            // FILTER 4: Location
            const filterLoc = document.getElementById("vhf-filter-location");
            if (filterLoc.value) {
                items = items.filter(function(i) { return i.loc === filterLoc.value; });
            }

            // FILTER 5: Month range - FIXED
            const filterMonth = document.getElementById("vhf-filter-month");
            if (filterMonth.value) {
                items = filterByMonthRange(items, filterMonth.value);
            }

            return items;
        }

        // Rendering Functions
        function renderStats(filtered) {
            document.getElementById("vhf-stat-total").textContent = filtered.length;
            
            var totalMinutes = filtered.reduce(function(sum, item) {
                return sum + item.highfives;
            }, 0);
            document.getElementById("vhf-stat-total-minutes").textContent = totalMinutes;

            var monthlyFiltered = filterCurrentMonth(allItems);
            // Only show current user's monthly data
            if (currentUserName) {
                monthlyFiltered = monthlyFiltered.filter(function(i) { 
                    return i.supervisor === currentUserName; 
                });
            }
            document.getElementById("vhf-stat-month").textContent = monthlyFiltered.length;
            document.getElementById("vhf-stat-month-label").textContent = new Date().toLocaleString("default", { month: "long", year: "numeric" });
        }
        
        function updateHeaders() {
            const suffix = dataMode === 'count' ? ' (Count)' : ' (Minutes)';
            const leaderboardText = dataMode === 'count' ? 'Total HighFives (Count)' : 'Total HighFives (min)';

            document.getElementById("vhf-chart-dept-title").textContent = 'High Fives by Department' + suffix;
            document.getElementById("vhf-chart-month-title").textContent = 'High Fives by Month' + suffix;
            document.getElementById("vhf-chart-location-title").textContent = 'High Fives by Location' + suffix;
            document.getElementById("vhf-employee-leaderboard-header").textContent = leaderboardText;
        }

        function renderTopEmployees(filtered) {
            var itemsToProcess = filtered;
            var totals = {};
            
            itemsToProcess.forEach(function(i) {
                if (!totals[i.employee]) {
                    totals[i.employee] = { value: 0, dept: i.dept };
                }
                totals[i.employee].value += dataMode === 'count' ? 1 : i.highfives;
                totals[i.employee].dept = i.dept;
            });

            var rows = Object.keys(totals).map(function(name) {
                return {
                    name: name,
                    value: totals[name].value,
                    dept: totals[name].dept
                };
            })
            .sort(function(a,b){ return b.value - a.value; })
            .slice(0, 10); // Show top 10 only

            document.getElementById("vhf-leaderboard-body").innerHTML = rows.map(function(r, idx) {
                return '<tr>' +
                    '<td>' + (idx+1) + '</td>' +
                    '<td>' + r.name + '</td>' +
                    '<td>' + r.value + '</td>' +
                    '<td>' + r.dept + '</td>' +
                '</tr>';
            }).join("");
        }

        function renderTable(filtered) {
            var recent = filtered.slice().sort(function(a,b){ return b.date - a.date; });

            document.getElementById("vhf-table-body").innerHTML = recent.map(function(r) {
                return '<tr>' +
                    '<td>' + r.date.toLocaleDateString() + '</td>' +
                    '<td>' + r.employee + '</td>' +
                    '<td>' + r.supervisor + '</td>' +
                    '<td>' + r.team + '</td>' +
                    '<td>' + r.dept + '</td>' +
                    '<td>' + r.loc + '</td>' +
                    '<td>' + r.reason.substring(0, 50) + (r.reason.length > 50 ? '...' : '') + '</td>' +
                    '<td>' + r.highfives + '</td>' +
                '</tr>';
            }).join("");
            
            var footerTotalMinutes = recent.reduce(function(sum, item) {
                return sum + item.highfives;
            }, 0);
            var footerTotalCount = recent.length;
            
            document.getElementById("vhf-footer-total-minutes").textContent = footerTotalCount + ' / ' + footerTotalMinutes;
        }
        
        function getChartColors(count) {
            const colors = [
                '#0078D4', '#E30008', '#00B294', '#FFB900', '#F27B42', 
                '#7719AA', '#008C95', '#C43D4B', '#99C74E', '#194A8A',
                '#004B87', '#5C2D91', '#4D4D4D', '#107C10', '#D83B01'
            ];
            let result = [];
            for (let i = 0; i < count; i++) {
                result.push(colors[i % colors.length]);
            }
            return result;
        }

        function renderCharts(filtered) {
            var deptAgg = {};
            var locAgg = {};
            var monthAgg = {};
            
            const aggValue = dataMode === 'count' ? 1 : 'highfives';

            filtered.forEach(function(i) {
                const value = aggValue === 1 ? 1 : i.highfives;

                if (!deptAgg[i.dept]) deptAgg[i.dept] = 0;
                deptAgg[i.dept] += value;

                if (!locAgg[i.loc]) locAgg[i.loc] = 0;
                locAgg[i.loc] += value;

                var ym = i.date.getFullYear() + "-" + String(i.date.getMonth()+1).padStart(2,"0");
                if (!monthAgg[ym]) monthAgg[ym] = 0;
                monthAgg[ym] += value;
            });
            
            var sortedMonthKeys = Object.keys(monthAgg).sort();
            var sortedMonthLabels = sortedMonthKeys.map(function(ym) {
                var parts = ym.split('-');
                var date = new Date(parts[0], parts[1] - 1);
                return date.toLocaleString('default', { month: 'short', year: 'numeric' });
            });
            var sortedMonthValues = sortedMonthKeys.map(function(key) { return monthAgg[key]; });

            if (chartDept) chartDept.destroy();
            if (chartLoc) chartLoc.destroy();
            if (chartMonth) chartMonth.destroy();
            
            const deptLabels = Object.keys(deptAgg);
            const deptValues = Object.values(deptAgg);

            chartDept = new Chart(document.getElementById("vhf-chart-dept"), {
                type: "doughnut",
                data: {
                    labels: deptLabels,
                    datasets: [{
                        data: deptValues,
                        backgroundColor: getChartColors(deptLabels.length)
                    }]
                },
                options: {
                    maintainAspectRatio: false,
                    cutout: "60%",
                    radius: '80%',
                    responsive: true,
                    plugins: {
                        legend: {
                            position: "right",
                            labels: {
                                font: { size: 12 },
                                usePointStyle: true
                            }
                        }
                    }
                }
            });
            
            const locLabels = Object.keys(locAgg);
            const locValues = Object.values(locAgg);

            chartLoc = new Chart(document.getElementById("vhf-chart-location"), {
                type: "bar",
                data: {
                    labels: locLabels,
                    datasets: [{
                        data: locValues,
                        backgroundColor: getChartColors(locLabels.length)[0],
                        borderColor: getChartColors(locLabels.length)[0],
                        borderWidth: 1
                    }]
                },
                options: {
                    maintainAspectRatio: false,
                    responsive: true,
                    plugins: { legend: { display: false } },
                    scales: {
                        y: { beginAtZero: true }
                    }
                }
            });

            chartMonth = new Chart(document.getElementById("vhf-chart-month"), {
                type: "line",
                data: {
                    labels: sortedMonthLabels,
                    datasets: [{
                        data: sortedMonthValues,
                        tension: 0.4,
                        backgroundColor: 'rgba(0, 120, 212, 0.2)',
                        borderColor: '#0078D4',
                        fill: true
                    }]
                },
                options: {
                    maintainAspectRatio: false,
                    responsive: true,
                    plugins: { legend: { display: false } },
                    scales: {
                        y: { beginAtZero: true }
                    }
                }
            });
        }

        function renderAll() {
            var filtered = applyFilters();
            updateHeaders();
            renderStats(filtered);
            renderTopEmployees(filtered);
            renderTable(filtered);
            renderCharts(filtered);
        }

        // Event Listeners
        function initEventListeners() {
            const filterClear = document.getElementById("vhf-filter-clear");
            const modeRadios = document.querySelectorAll('input[name="data-mode"]');
            
            if (filterClear) {
                filterClear.addEventListener("click", function() {
                    document.getElementById("vhf-filter-search").value = "";
                    document.getElementById("vhf-filter-dept").value = "";
                    document.getElementById("vhf-filter-location").value = "";
                    document.getElementById("vhf-filter-month").value = "";
                    renderAll();
                });
            }

            const filterElements = [
                "vhf-filter-search", "vhf-filter-dept", "vhf-filter-location", "vhf-filter-month"
            ];
            
            filterElements.forEach(function(id) {
                const el = document.getElementById(id);
                if (el) {
                    el.addEventListener("change", renderAll);
                    el.addEventListener("input", renderAll);
                }
            });
            
            if (modeRadios.length > 0) {
                modeRadios.forEach(function(radio) {
                    radio.addEventListener('change', function() {
                        dataMode = this.value;
                        renderAll();
                    });
                });
            }
        }

        // ====== INITIALIZATION ======
        async function initializeApp() {
            console.log("Initializing Virtual High Five Dashboard...");
            
            try {
                // Show random quote immediately
                showRandomQuote();
                
                // Get current user first
                await getCurrentUser();
                
                // Set today's date in form
                const formDate = document.getElementById('form-date');
                if (formDate) {
                    formDate.valueAsDate = new Date();
                }
                
                // Initialize form components
                await syncFormChoices();
                initPeoplePicker();
                initFormSubmission();
                
                // Initialize dashboard
                await loadData();
                initEventListeners();
                
                console.log("Dashboard initialized successfully");
                
            } catch (error) {
                console.error("Initialization error:", error);
                showStatus("Error initializing dashboard. Please refresh the page.", true);
            }
        }

        // Start the application when DOM is loaded
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializeApp);
        } else {
            initializeApp();
        }