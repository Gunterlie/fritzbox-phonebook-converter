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
    contacts.forEach(c => addRow(c.name, c.home, c.mobile, c.work));
  }
}

function addRow(name = '', home = '', mobile = '', work = '') {
  const body = document.getElementById('contacts-body');
  const tr = document.createElement('tr');

  const makeInput = (className, value) => {
    const input = document.createElement('input');
    input.type = 'text';
    input.className = className;
    input.value = value;
    const td = document.createElement('td');
    td.appendChild(input);
    return td;
  };

  tr.appendChild(makeInput('name', name));
  tr.appendChild(makeInput('home', home));
  tr.appendChild(makeInput('mobile', mobile));
  tr.appendChild(makeInput('work', work));

  const actionTd = document.createElement('td');
  actionTd.className = 'row-actions';
  const removeBtn = document.createElement('button');
  removeBtn.textContent = '✕';
  removeBtn.onclick = () => tr.remove();
  actionTd.appendChild(removeBtn);
  tr.appendChild(actionTd);

  body.appendChild(tr);
}

function collectContacts() {
  const rows = document.querySelectorAll('#contacts-body tr');
  const contacts = [];
  rows.forEach(row => {
    const name = row.querySelector('.name').value.trim();
    const home = row.querySelector('.home').value.trim();
    const mobile = row.querySelector('.mobile').value.trim();
    const work = row.querySelector('.work').value.trim();
    if (name || home || mobile || work) contacts.push({ name, home, mobile, work });
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
    { header: 'Telefon (privat)', key: 'home', width: 22 },
    { header: 'Mobil', key: 'mobile', width: 22 },
    { header: 'Geschäftlich', key: 'work', width: 22 },
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
    const row = ws.addRow({ name: c.name, home: c.home, mobile: c.mobile, work: c.work });
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

  ws.autoFilter = { from: 'A1', to: 'D1' };

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
      rows.push([values[1], values[2], values[3], values[4]]);
    });

    let dataRows = rows;
    if (rows.length > 0) {
      const first = rows[0];
      const looksLikeHeader = String(first[0] || '').trim().toLowerCase() === 'name';
      if (looksLikeHeader) dataRows = rows.slice(1);
    }

    const contacts = dataRows
      .map(r => ({
        name: String(r[0] || '').trim(),
        home: String(r[1] || '').trim(),
        mobile: String(r[2] || '').trim(),
        work: String(r[3] || '').trim(),
      }))
      .filter(c => c.name || c.home || c.mobile || c.work);

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
    .filter(c => c.name || c.home || c.mobile || c.work)
    .map((c, idx) => {
      const numbers = [
        c.home && { type: 'home', value: c.home },
        c.mobile && { type: 'mobile', value: c.mobile },
        c.work && { type: 'work', value: c.work },
      ].filter(Boolean);

      const numbersXml = numbers
        .map((n, numIdx) => `        <number type="${n.type}" prio="${numIdx === 0 ? '1' : '0'}" id="${numIdx}">${escapeXml(n.value)}</number>`)
        .join('\n');

      return `    <contact>
      <category>0</category>
      <person>
        <realName>${escapeXml(c.name)}</realName>
      </person>
      <telephony nid="${numbers.length}">
${numbersXml}
      </telephony>
      <services />
      <setup />
      <features doorphone="0" />
      <uniqueid>${idx + 1}</uniqueid>
    </contact>`;
    })
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
