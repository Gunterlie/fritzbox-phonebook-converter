const STORAGE_KEY = 'fritzbox-contacts';

function loadContacts() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function renderContacts(contacts) {
  const body = document.getElementById('contacts-body');
  body.innerHTML = '';
  if (contacts.length === 0) {
    addRow();
  } else {
    contacts.forEach(c => addRow(c.name, c.number));
  }
}

function addRow(name = '', number = '') {
  const body = document.getElementById('contacts-body');
  const tr = document.createElement('tr');
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.className = 'name';
  nameInput.value = name;
  const numberInput = document.createElement('input');
  numberInput.type = 'text';
  numberInput.className = 'number';
  numberInput.value = number;

  const nameTd = document.createElement('td');
  nameTd.appendChild(nameInput);
  const numberTd = document.createElement('td');
  numberTd.appendChild(numberInput);
  const actionTd = document.createElement('td');
  actionTd.className = 'row-actions';
  const removeBtn = document.createElement('button');
  removeBtn.textContent = '✕';
  removeBtn.onclick = () => tr.remove();
  actionTd.appendChild(removeBtn);

  tr.appendChild(nameTd);
  tr.appendChild(numberTd);
  tr.appendChild(actionTd);
  body.appendChild(tr);
}

function collectContacts() {
  const rows = document.querySelectorAll('#contacts-body tr');
  const contacts = [];
  rows.forEach(row => {
    const name = row.querySelector('.name').value.trim();
    const number = row.querySelector('.number').value.trim();
    if (name || number) contacts.push({ name, number });
  });
  return contacts;
}

function showStatus(text) {
  const status = document.getElementById('status');
  status.textContent = text;
  setTimeout(() => { status.textContent = ''; }, 3000);
}

function saveContacts() {
  const contacts = collectContacts();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(contacts));
  showStatus(`Gespeichert (${contacts.length} Einträge)`);
}

async function downloadExcel() {
  const contacts = collectContacts();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Fritzbox Telefonbuch Generator';
  wb.created = new Date();

  const ws = wb.addWorksheet('Telefonbuch', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  ws.columns = [
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Telefonnummer', key: 'number', width: 22 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });
  headerRow.height = 20;

  contacts.forEach((c, idx) => {
    const row = ws.addRow({ name: c.name, number: c.number });
    row.eachCell(cell => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFDDDDDD' } },
        left: { style: 'thin', color: { argb: 'FFDDDDDD' } },
        bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
        right: { style: 'thin', color: { argb: 'FFDDDDDD' } },
      };
      if (idx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F6FC' } };
      }
    });
  });

  ws.autoFilter = { from: 'A1', to: 'B1' };

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'telefonbuch.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function uploadExcel(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async (e) => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(e.target.result);
    const ws = wb.worksheets[0];

    const rows = [];
    ws.eachRow((row) => {
      const values = row.values;
      rows.push([values[1], values[2]]);
    });

    let dataRows = rows;
    if (rows.length > 0) {
      const first = rows[0];
      const looksLikeHeader = String(first[0] || '').trim().toLowerCase() === 'name';
      if (looksLikeHeader) dataRows = rows.slice(1);
    }

    const contacts = dataRows
      .map(r => ({ name: String(r[0] || '').trim(), number: String(r[1] || '').trim() }))
      .filter(c => c.name || c.number);

    renderContacts(contacts);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(contacts));
    showStatus('Excel-Datei geladen');
  };
  reader.readAsArrayBuffer(file);
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildFritzboxXml(contacts) {
  const contactsXml = contacts
    .filter(c => c.name || c.number)
    .map((c, idx) => `    <contact>
      <category>0</category>
      <person>
        <realName>${escapeXml(c.name)}</realName>
      </person>
      <telephony nid="1">
        <number type="home" prio="1" id="${idx}">${escapeXml(c.number)}</number>
      </telephony>
      <services />
      <setup />
      <features doorphone="0" />
      <uniqueid>${idx + 1}</uniqueid>
    </contact>`)
    .join('\n');

  return `<?xml version="1.0" encoding="utf-8"?>
<phonebooks>
  <phonebook name="Telefonbuch">
${contactsXml}
  </phonebook>
</phonebooks>
`;
}

function downloadXml() {
  const contacts = collectContacts();
  const xml = buildFritzboxXml(contacts);
  const blob = new Blob([xml], { type: 'application/xml' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'fritzbox_telefonbuch.xml';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

renderContacts(loadContacts());
