let tempoData = null;
let timeLaborData = null;
let trcValues = new Set();
let selectedTrcValues = new Set();

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('tempoFile').addEventListener('change', function(e) {
        updateFileStatus('tempoStatus', e.target.files[0]);
    });
    
    document.getElementById('timeLaborFile').addEventListener('change', function(e) {
        updateFileStatus('timeLaborStatus', e.target.files[0]);
    });
    
    document.getElementById('processButton').addEventListener('click', processExcelFiles);
    
});

function updateFileStatus(statusId, file) {
    const statusElement = document.getElementById(statusId);
    if (file) {
        statusElement.textContent = `Selected: ${file.name} (${formatFileSize(file.size)})`;
    } else {
        statusElement.textContent = 'No file selected';
    }
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' bytes';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    cellDates: true,
                    cellStyles: true,
                    raw: false
                });
                resolve(workbook);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = function() {
            reject(new Error('Error reading file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

function processExcelFiles() {
    const tempoFile = document.getElementById('tempoFile').files[0];
    const timeLaborFile = document.getElementById('timeLaborFile').files[0];
    const output = document.getElementById('output');
    
    if (!tempoFile || !timeLaborFile) {
        output.innerHTML = '<p style="color: red;">Please select both files before processing.</p>';
        return;
    }
    
    output.innerHTML = '<p>Processing files, please wait...</p>';
    
    Promise.all([
        readExcelFile(tempoFile),
        readExcelFile(timeLaborFile)
    ])
    .then(([tempo, timeLabor]) => {
        tempoData = tempo;
        timeLaborData = timeLabor;
        
        compareEmployeeData();
    })
    .catch(error => {
        output.innerHTML = `<p style="color: red;">Error processing files: ${error.message}</p>`;
    });
}

function findHeaderRow(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    for (let r = range.s.r; r < Math.min(range.s.r + 10, range.e.r); ++r) {
        const potentialHeaders = [];
        let headerCandidates = 0;
        
        for (let c = range.s.c; c <= range.e.c; ++c) {
            const cellRef = XLSX.utils.encode_cell({r: r, c: c});
            const cell = sheet[cellRef];
            
            if (cell && cell.v) {
                potentialHeaders.push(String(cell.v).toLowerCase());
                
                const headerText = String(cell.v).toLowerCase();
                if (headerText.includes('name') || 
                    headerText.includes('id') || 
                    headerText.includes('hours') || 
                    headerText.includes('employee') ||
                    headerText.includes('date') ||
                    headerText.includes('task') ||
                    headerText.includes('project')) {
                    headerCandidates++;
                }
            }
        }
        
        if (headerCandidates >= 3) {
            return r;
        }
    }
    
    return range.s.r;
}

function processTempoData() {
    if (!tempoData) {
        return null;
    }
    
    const sheetName = tempoData.SheetNames[0];
    const worksheet = tempoData.Sheets[sheetName];
    
    const headerRowIndex = findHeaderRow(worksheet);
    
    const rawData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false,
        defval: ''
    });
    
    if (rawData.length <= headerRowIndex) {
        return null;
    }
    
    const headers = rawData[headerRowIndex].map(h => String(h || '').toLowerCase());
    
    const nameIndex = headers.findIndex(h => h.includes('name'));
    const employeeIdIndex = headers.findIndex(h => h.includes('id') && (h.includes('employee') || h.includes('empl')));
    const hoursIndex = headers.findIndex(h => h.includes('hours'));
    const taskIndex = headers.findIndex(h => h.includes('task'));
    
    if (nameIndex === -1 || employeeIdIndex === -1 || hoursIndex === -1) {
        if (nameIndex === -1 && headers.length > 0) nameIndex = 0;
        if (employeeIdIndex === -1 && headers.length > 1) employeeIdIndex = 1;
        if (hoursIndex === -1 && headers.length > 6) hoursIndex = 6;
        if (taskIndex === -1 && headers.length > 5) taskIndex = 5;
    }
    
    if (nameIndex === -1 || employeeIdIndex === -1 || hoursIndex === -1) {
        return null;
    }
    
    const employeeHoursMap = new Map();
    let admFreeDaysCount = 0;
    
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        
        if (!row || row.length === 0) continue;
        
        const name = row[nameIndex] || "";
        const employeeId = row[employeeIdIndex] ? String(row[employeeIdIndex]) : "";
        
        let hours = 0;
        if (row[hoursIndex]) {
            const hoursStr = String(row[hoursIndex]).replace(',', '.');
            hours = parseFloat(hoursStr) || 0;
        }
        
        const task = taskIndex !== -1 ? (row[taskIndex] || "") : "";
        if (task.toLowerCase && task.toLowerCase().includes("adm free days")) {
            admFreeDaysCount++;
            continue;
        }
        
        if (!employeeId || employeeId === "" || name.toUpperCase() === "TOTAL") {
            continue;
        }
        
        if (employeeHoursMap.has(employeeId)) {
            const employee = employeeHoursMap.get(employeeId);
            employee.totalHours += hours;
        } else {
            employeeHoursMap.set(employeeId, {
                id: employeeId,
                name: name,
                totalHours: hours
            });
        }
    }
    
    const employees = Array.from(employeeHoursMap.values());
    
    return {
        employees: employees,
        admFreeDaysCount: admFreeDaysCount
    };
}

function processTimeLaborData() {
    if (!timeLaborData) {
        return null;
    }
    
    const sheetName = timeLaborData.SheetNames[0];
    const worksheet = timeLaborData.Sheets[sheetName];
    
    const headerRowIndex = findHeaderRow(worksheet);
    
    const rawData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false,
        defval: ''
    });
    
    if (rawData.length <= headerRowIndex) {
        return null;
    }
    
    const headers = rawData[headerRowIndex].map(h => String(h || '').toLowerCase());
    
    const nameIndex = headers.findIndex(h => h.includes('name') && (h.includes('employee') || h.includes('empl')));
    const employeeIdIndex = headers.findIndex(h => h.includes('id') && (h.includes('empl')));
    const hoursIndex = headers.findIndex(h => h.includes('hours'));
    const statusIndex = headers.findIndex(h => h.includes('status'));
    const trcIndex = headers.findIndex(h => h.includes('trc') && h.includes('desc'));
    
    if (nameIndex === -1 || employeeIdIndex === -1 || hoursIndex === -1) {
        if (nameIndex === -1 && headers.length > 0) nameIndex = 0;
        if (employeeIdIndex === -1 && headers.length > 1) employeeIdIndex = 1;
        if (hoursIndex === -1 && headers.length > 5) hoursIndex = 5;
        if (statusIndex === -1 && headers.length > 6) statusIndex = 6;
    }
    
    if (nameIndex === -1 || employeeIdIndex === -1 || hoursIndex === -1) {
        return null;
    }
    
    const employeeMap = new Map();
    const newTrcValues = new Set();
    
    // First pass: collect all TRC values
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        const trc = trcIndex !== -1 ? (row[trcIndex] || "").toString().trim() : "";
        if (trc) {
            newTrcValues.add(trc);
        }
    }

    // Update global TRC sets
    trcValues = newTrcValues;
    // Initialize selectedTrcValues as empty set if not already set
    if (!selectedTrcValues) {
        selectedTrcValues = new Set();
    }
    
    // Second pass: process employee data
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        
        if (!row || row.length === 0) continue;
        
        const name = row[nameIndex] || "";
        const employeeId = row[employeeIdIndex] ? String(row[employeeIdIndex]) : "";
        const trc = trcIndex !== -1 ? (row[trcIndex] || "").toString().trim() : "";
        
        let hours = 0;
        if (row[hoursIndex]) {
            const hoursStr = String(row[hoursIndex]).replace(',', '.');
            hours = parseFloat(hoursStr) || 0;
        }
        
        const status = statusIndex !== -1 ? (row[statusIndex] || 'Not Available') : 'Not Available';
        
        if (!employeeId || employeeId === "" || name.toUpperCase() === "TOTAL") {
            continue;
        }
        
        if (employeeMap.has(employeeId)) {
            const existingEmployee = employeeMap.get(employeeId);
            if (!existingEmployee.trcHours) {
                existingEmployee.trcHours = new Map();
            }
            existingEmployee.trcHours.set(trc, (existingEmployee.trcHours.get(trc) || 0) + hours);
            if (selectedTrcValues.has(trc)) {
                existingEmployee.hours += hours;
            }
        } else {
            const trcHours = new Map();
            trcHours.set(trc, hours);
            employeeMap.set(employeeId, {
                id: employeeId,
                name: name,
                hours: selectedTrcValues.has(trc) ? hours : 0,
                trcHours: trcHours,
                validator: status,
                validationTime: 'Not Available'
            });
        }
    }
    
    return employeeMap;
}

function compareEmployeeData() {
    const tempoResult = processTempoData();
    const timeLaborEmployees = processTimeLaborData();
    
    if (!tempoResult) {
        document.getElementById('output').innerHTML = 
            '<p style="color: red;">Error: Unable to process Tempo file.</p>';
        return;
    }
    
    const tempoEmployees = tempoResult.employees;
    const admFreeDaysCount = tempoResult.admFreeDaysCount;
    
    if (!tempoEmployees || !timeLaborEmployees) {
        document.getElementById('output').innerHTML = 
            '<p style="color: red;">Error: Unable to process one or both files.</p>';
        return;
    }
    
    const discrepancies = [];
    
    tempoEmployees.forEach(tempoEmployee => {
        const employeeId = tempoEmployee.id;
        const timeLaborEmployee = timeLaborEmployees.get(employeeId);
        
        if (!employeeId || !tempoEmployee.name) {
            return;
        }
        
        if (!timeLaborEmployee) {
            discrepancies.push({
                id: employeeId,
                name: tempoEmployee.name,
                tempoHours: tempoEmployee.totalHours,
                timeLaborHours: 0,
                status: "Missing in Time & Labor",
                validator: "N/A",
                validationTime: "N/A"
            });
        }
        else if (Math.abs(tempoEmployee.totalHours - timeLaborEmployee.hours) > 0.01) {
            discrepancies.push({
                id: employeeId,
                name: tempoEmployee.name,
                tempoHours: tempoEmployee.totalHours,
                timeLaborHours: timeLaborEmployee.hours,
                status: "Hours Mismatch",
                validator: timeLaborEmployee.validator,
                validationTime: timeLaborEmployee.validationTime
            });
        }
    });
    
    timeLaborEmployees.forEach((timeLaborEmployee, employeeId) => {
        const tempoEmployee = tempoEmployees.find(e => e.id === employeeId);
        if (!tempoEmployee) {
            discrepancies.push({
                id: employeeId,
                name: timeLaborEmployee.name,
                tempoHours: 0,
                timeLaborHours: timeLaborEmployee.hours,
                status: "Missing in Tempo",
                validator: timeLaborEmployee.validator,
                validationTime: timeLaborEmployee.validationTime
            });
        }
    });
    
    displayDiscrepancies(discrepancies, admFreeDaysCount, tempoResult, timeLaborEmployees);
}

function updateTimeLaborHours() {
    if (!tempoData || !timeLaborData) {
        return;
    }
    compareEmployeeData();
}

function displayDiscrepancies(discrepancies, admFreeDaysCount, tempoResult, timeLaborEmployees) {
    const output = document.getElementById('output');
    
    const allEmployees = new Map();
    
    tempoResult.employees.forEach(emp => {
        allEmployees.set(emp.id, {
            id: emp.id,
            name: emp.name,
            tempoHours: emp.totalHours,
            timeLaborHours: 0,
            trcHours: new Map(),
            status: "OK",
            validator: "N/A",
            validationTime: "N/A"
        });
    });
    
    timeLaborEmployees.forEach((emp, id) => {
        if (allEmployees.has(id)) {
            const employee = allEmployees.get(id);
            employee.timeLaborHours = emp.hours;
            employee.trcHours = emp.trcHours;
            employee.validator = emp.validator;
            employee.validationTime = emp.validationTime;
            
            if (Math.abs(employee.tempoHours - emp.hours) > 0.01) {
                employee.status = "Hours Mismatch";
            }
        } else {
            allEmployees.set(id, {
                id: id,
                name: emp.name,
                tempoHours: 0,
                timeLaborHours: emp.hours,
                trcHours: emp.trcHours,
                status: "Missing in Tempo",
                validator: emp.validator,
                validationTime: emp.validationTime
            });
        }
    });
    
    allEmployees.forEach(emp => {
        if (emp.timeLaborHours === 0 && emp.tempoHours > 0) {
            emp.status = "Missing in Time & Labor";
        }
    });
    
    const allEmployeesArray = Array.from(allEmployees.values());
    const discrepancyCount = allEmployeesArray.filter(emp => emp.status !== "OK").length;
    
    allEmployeesArray.sort((a, b) => {
        if (a.status !== "OK" && b.status === "OK") return -1;
        if (a.status === "OK" && b.status !== "OK") return 1;
        
        if (a.status !== "OK" && b.status !== "OK") {
            const statusOrder = {
                "Hours Mismatch": 1,
                "Missing in Time & Labor": 2,
                "Missing in Tempo": 3
            };
            if (statusOrder[a.status] !== statusOrder[b.status]) {
                return statusOrder[a.status] - statusOrder[b.status];
            }
        }
        
        return a.name.localeCompare(b.name);
    });

    const trcCheckboxesHtml = Array.from(trcValues)
        .sort()
        .map(trc => {
            let totalHours = 0;
            allEmployeesArray.forEach(emp => {
                if (emp.trcHours.has(trc)) {
                    totalHours += emp.trcHours.get(trc);
                }
            });
            return `
                <div style="margin: 5px 0; flex: 0 0 calc(25% - 5px); white-space: nowrap;">
                    <input type="checkbox" id="trc_${trc}" class="trc-checkbox" value="${trc}" ${selectedTrcValues.has(trc) ? 'checked' : ''}>
                    <label for="trc_${trc}">${trc}</label>
                    <span style="color: #666; font-size: 0.9em; margin-left: 5px;">(${totalHours.toFixed(2)}h)</span>
                </div>
            `;
        }).join('');
    
    let html = `
        <h2>Employee Hours Report</h2>
        ${discrepancyCount === 0 
            ? '<div class="alert alert-success">' +
              '<i class="fas fa-check-circle"></i>' +
              '<span>Perfect match! No discrepancies found between Tempo and Time & Labor data.</span>' +
              '</div>'
            : '<div class="alert alert-warning">' +
              '<i class="fas fa-exclamation-triangle"></i>' +
              `<span>Found ${discrepancyCount} discrepancies between Tempo and Time & Labor data.</span>` +
              '</div>'
        }
        ${admFreeDaysCount > 0 ? `
            <div class="alert" style="background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1;">
                <i class="fas fa-info-circle"></i>
                <span>${admFreeDaysCount} entries with "ADM Free days" were excluded from the comparison.</span>
            </div>` : ''}
        <div class="trc-filters">
            <div class="checkbox-container">
                <input type="checkbox" id="showAllEmployees" checked>
                <label for="showAllEmployees">Show all employees (uncheck to show only discrepancies)</label>
            </div>
            <div>
                <div style="margin-bottom: 1rem;">
                    <strong>TRC Filters</strong>
                    <button id="selectAllTrc" style="margin-left: 1rem;">
                        <i class="fas fa-check-square"></i> Select All
                    </button>
                    <button id="deselectAllTrc" style="margin-left: 0.5rem;">
                        <i class="fas fa-square"></i> Deselect All
                    </button>
                </div>
                <div style="border: 1px solid #e2e8f0; padding: 1rem; border-radius: 8px;">
                    <div style="display: flex; flex-wrap: wrap; gap: 5px;">
                        ${trcCheckboxesHtml}
                    </div>
                </div>
            </div>
        </div>
        <table>
            <thead>
                <tr>
                    <th>Employee ID</th>
                    <th>Name</th>
                    <th style="text-align: right;">Tempo Hours</th>
                    <th style="text-align: right;">Time & Labor Hours</th>
                    <th style="text-align: right;">Difference</th>
                </tr>
            </thead>
            <tbody id="employeeTableBody">
    `;
    
    if (allEmployeesArray.length === 0) {
        html += `
            <tr>
                <td colspan="5" style="text-align: center; padding: 2rem; color: #64748b;">
                    <i class="fas fa-folder-open" style="font-size: 2rem; margin-bottom: 1rem; display: block;"></i>
                    No employees found in either system.
                </td>
            </tr>
        `;
    } else {
        allEmployeesArray.forEach(employee => {
            const difference = employee.tempoHours - employee.timeLaborHours;
            const differenceClass = difference > 0 ? 'positive' : (difference < 0 ? 'negative' : '');
            
            let rowStyle = '';
            let rowClass = employee.status === 'OK' ? 'match-row' : 'discrepancy-row';
            
            switch(employee.status) {
                case 'Missing in Time & Labor':
                    rowStyle = 'background-color: #fff7ed;';
                    break;
                case 'Missing in Tempo':
                    rowStyle = 'background-color: #f0f9ff;';
                    break;
            }
            
            html += `
                <tr class="${rowClass}" style="${rowStyle}">
                    <td>${employee.id}</td>
                    <td>${employee.name}</td>
                    <td style="text-align: right; font-variant-numeric: tabular-nums;">${employee.tempoHours.toFixed(2)}</td>
                    <td style="text-align: right; font-variant-numeric: tabular-nums;">${employee.timeLaborHours.toFixed(2)}</td>
                    <td style="text-align: right; font-variant-numeric: tabular-nums; ${differenceClass === 'positive' ? 'color: #dc2626;' : 
                                                   (differenceClass === 'negative' ? 'color: #2563eb;' : '')}">
                        ${difference.toFixed(2)}
                    </td>
                </tr>
            `;
        });
    }
    
    html += `
            </tbody>
        </table>
        <div class="legend">
            <p style="font-weight: 600; color: #334155; margin-bottom: 1rem;">Legend</p>
            <ul>
                <li><i class="fas fa-check" style="color: #10b981;"></i> Hours match between systems</li>
                <li><i class="fas fa-exclamation-triangle" style="color: #f97316;"></i> Missing in Time & Labor</li>
                <li><i class="fas fa-question-circle" style="color: #3b82f6;"></i> Missing in Tempo</li>
                <li><i class="fas fa-not-equal" style="color: #ef4444;"></i> Hours Mismatch</li>
                <li><span style="color: #dc2626;">Positive difference</span> More hours in Tempo</li>
                <li><span style="color: #2563eb;">Negative difference</span> More hours in Time & Labor</li>
                <li><span style="background: #fff7ed; padding: 2px 8px; border-radius: 4px;">Orange background</span> Missing in Time & Labor</li>
                <li><span style="background: #f0f9ff; padding: 2px 8px; border-radius: 4px;">Blue background</span> Missing in Tempo</li>
            </ul>
        </div>
    `;
    
    output.innerHTML = html;
    
    document.getElementById('showAllEmployees').addEventListener('change', function() {
        const matchRows = document.querySelectorAll('.match-row');
        matchRows.forEach(row => {
            row.style.display = this.checked ? '' : 'none';
        });
    });

    document.getElementById('selectAllTrc').addEventListener('click', function() {
        selectedTrcValues = new Set(trcValues);
        document.querySelectorAll('.trc-checkbox').forEach(checkbox => {
            checkbox.checked = true;
        });
        updateTimeLaborHours();
    });

    document.getElementById('deselectAllTrc').addEventListener('click', function() {
        selectedTrcValues = new Set();  // Create a new empty set
        document.querySelectorAll('.trc-checkbox').forEach(checkbox => {
            checkbox.checked = false;
        });
        updateTimeLaborHours();
    });

    document.querySelectorAll('.trc-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', function(event) {
            // Prevent the default behavior
            event.preventDefault();
            
            if (this.checked) {
                selectedTrcValues.add(this.value);
            } else {
                selectedTrcValues.delete(this.value);
            }
            
            // Manually update the checkbox state
            this.checked = selectedTrcValues.has(this.value);
            
            updateTimeLaborHours();
        });
    });
}