document.getElementById('crmForm').addEventListener('submit', function(event) {
    event.preventDefault();

    let clientName = document.getElementById('clientName').value;
    let contactPerson = document.getElementById('contactPerson').value;
    let contactNumber = document.getElementById('contactNumber').value;
    let contactMethod = document.getElementById('contactMethod').value;
    let clientStatus = document.getElementById('clientStatus').value;
    let interestLevel = document.getElementById('interestLevel').value;
    let communicationResult = document.getElementById('communicationResult').value;
    let notes = document.getElementById('notes').value;
    let salesRep = document.getElementById('salesRep').value;
    let employeeCode = document.getElementById('employeeCode').value;
    let branch = document.getElementById('branch').value;

    let client = {
        id: Date.now(),
        clientName,
        contactPerson,
        contactNumber,
        contactMethod,
        clientStatus,
        interestLevel,
        communicationResult,
        notes,
        salesRep,
        employeeCode,
        branch
    };

    saveClient(client);

    let row = createTableRow(client);

    document.getElementById('clientList').appendChild(row);

    document.getElementById('crmForm').reset();
});

function saveClient(client) {
    let clients = JSON.parse(localStorage.getItem('clients')) || [];
    clients.push(client);
    localStorage.setItem('clients', JSON.stringify(clients));
}

function loadClients() {
    let clients = JSON.parse(localStorage.getItem('clients')) || [];
    clients.forEach(client => {
        let row = createTableRow(client);
        document.getElementById('clientList').appendChild(row);
    });
}

function createTableRow(client) {
    let row = document.createElement('tr');

    row.insertCell(0).textContent = client.clientName;
    row.insertCell(1).textContent = client.contactPerson;
    row.insertCell(2).textContent = client.contactNumber;
    row.insertCell(3).textContent = client.contactMethod;
    row.insertCell(4).textContent = client.clientStatus;
    row.insertCell(5).textContent = client.interestLevel;
    row.insertCell(6).textContent = client.communicationResult;
    row.insertCell(7).textContent = client.notes;
    row.insertCell(8).textContent = client.salesRep;
    row.insertCell(9).textContent = client.employeeCode;
    row.insertCell(10).textContent = client.branch;

    let deleteCell = row.insertCell(11);
    let deleteBtn = document.createElement('button');
    deleteBtn.textContent = 'حذف';
    deleteBtn.className = 'btn btn-danger btn-sm';
    deleteBtn.onclick = function() {
        deleteClient(client.id, row);
    };
    deleteCell.appendChild(deleteBtn);

    return row;
}

function deleteClient(id, row) {
    let clients = JSON.parse(localStorage.getItem('clients')) || [];
    clients = clients.filter(client => client.id !== id);
    localStorage.setItem('clients', JSON.stringify(clients));
    row.remove();
}

document.getElementById('exportBtn').addEventListener('click', function() {
    let table = document.querySelector('table');
    let wb = XLSX.utils.table_to_book(table, {sheet: "CRM Data"});
    XLSX.writeFile(wb, 'CRM_Data.xlsx');
});

window.addEventListener('load', loadClients);
