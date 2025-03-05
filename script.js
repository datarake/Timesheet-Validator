let tempoData = null;
let timeLaborData = null;

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
                    cellStyles: true
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

function processTempoData() {
    if (!tempoData) {
        return null;
    }
    
    const sheetName = tempoData.SheetNames[0];
    const worksheet = tempoData.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    const employeeHoursMap = new Map();
    
    jsonData.forEach(row => {
        const name = row['Team Activity Details'] || "";
        const employeeId = row['__EMPTY'] ? row['__EMPTY'].toString() : "";
        const hours = parseFloat(row['__EMPTY_6']) || 0;
        
        if (!employeeId || employeeId === "" || name === "TOTAL") {
            return;
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
    });
    
    return Array.from(employeeHoursMap.values());
}

function processTimeLaborData() {
    if (!timeLaborData) {
        return null;
    }
    
    const sheetName = timeLaborData.SheetNames[0];
    const worksheet = timeLaborData.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    const employeeMap = new Map();
    
    jsonData.forEach(row => {
        let employeeId = "";
        if (row['Timesheet Summary RPTD MGR']) {
            employeeId = row['Timesheet Summary RPTD MGR'].toString();
        }
        
        let name = "";
        for (const key in row) {
            if (!isNaN(parseInt(key)) && typeof row[key] === 'string' && 
                row[key].includes(',')) {
                name = row[key];
                break;
            }
        }
        
        if (!employeeId || employeeId === "" || name === "TOTAL") {
            return;
        }
        
        let hours = 0;
        if (row['__EMPTY_3'] !== undefined) {
            hours = parseFloat(row['__EMPTY_3']) || 0;
        }
        
        if (employeeId) {
            employeeMap.set(employeeId, {
                id: employeeId,
                name: name,
                hours: hours
            });
        }
    });
    
    return employeeMap;
}

function compareEmployeeData() {
    const tempoEmployees = processTempoData();
    const timeLaborEmployees = processTimeLaborData();
    
    if (!tempoEmployees || !timeLaborEmployees) {
        document.getElementById('output').innerHTML = 
            '<p style="color: red;">Error: Unable to process one or both files.</p>';
        return;
    }
    
    const discrepancies = [];
    
    tempoEmployees.forEach(tempoEmployee => {
        const employeeId = tempoEmployee.id;
        const timeLaborEmployee = timeLaborEmployees.get(employeeId);
        
        if (employeeId === "EmployeeId" || tempoEmployee.name === "Name" || 
            !employeeId || !tempoEmployee.name) {
            return;
        }
        
        if (tempoEmployee.name === "TOTAL") {
            return;
        }
        
        if (timeLaborEmployee && Math.abs(tempoEmployee.totalHours - timeLaborEmployee.hours) > 0.01) {
            discrepancies.push({
                id: employeeId,
                name: tempoEmployee.name,
                tempoHours: tempoEmployee.totalHours,
                timeLaborHours: timeLaborEmployee.hours,
                status: "Hours Mismatch"
            });
        }
    });
    
    displayDiscrepancies(discrepancies);
}

function displayDiscrepancies(discrepancies) {
    const output = document.getElementById('output');
    
    if (!discrepancies || discrepancies.length === 0) {
        output.innerHTML = '<p>No hours mismatches found between files.</p>';
        return;
    }
    
    const filteredDiscrepancies = discrepancies.filter(emp => 
        emp.id !== "EmployeeId" && 
        emp.name !== "Name" && 
        emp.name !== "TOTAL" &&
        emp.status !== "Missing in Time & Labor"
    );
    
    if (filteredDiscrepancies.length === 0) {
        output.innerHTML = '<p>No hours mismatches found between files.</p>';
        return;
    }
    
    let html = `
        <h2>Employee Discrepancies</h2>
        <p>Found ${filteredDiscrepancies.length} employees with hour mismatches between Tempo and Time & Labor data.</p>
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th>Employee ID</th>
                    <th>Name</th>
                    <th>Tempo Hours</th>
                    <th>Time & Labor Hours</th>
                    <th>Difference</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    filteredDiscrepancies.forEach(employee => {
        const difference = employee.tempoHours - employee.timeLaborHours;
        const differenceClass = difference > 0 ? 'positive' : (difference < 0 ? 'negative' : '');
        
        html += `
            <tr>
                <td>${employee.id}</td>
                <td>${employee.name}</td>
                <td style="text-align: right;">${employee.tempoHours}</td>
                <td style="text-align: right;">${employee.timeLaborHours}</td>
                <td style="text-align: right; ${differenceClass === 'positive' ? 'color: red;' : 
                                               (differenceClass === 'negative' ? 'color: blue;' : '')}">
                    ${difference.toFixed(2)}
                </td>
                <td>${employee.status}</td>
            </tr>
        `;
    });
    
    html += `
            </tbody>
        </table>
        <p><strong>Legend:</strong> <span style="color: red;">Positive difference</span> - More hours in Tempo than in Time & Labor, 
        <span style="color: blue;">Negative difference</span> - More hours in Time & Labor than in Tempo</p>
    `;
    
    output.innerHTML = html;
}