document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('registerFile')
  const fileName = document.getElementById('registerFileName')
  const status = document.getElementById('registerStatus')
  const processBtn = document.getElementById('processRegisterBtn')
  const summary = document.getElementById('summaryText')
  const tableBody = document.getElementById('tableBody')
  const clientSelect = document.getElementById('registerClient')
  const processingDots = document.getElementById('registerProcessingDots')
  const cancelRowEditBtn = document.getElementById('cancelRowEditBtn')
  const saveRowEditBtn = document.getElementById('saveRowEditBtn')

  let capturedRows = []


  if (cancelRowEditBtn) {
    cancelRowEditBtn.addEventListener('click', function () {
      closeRowModal()
    })
  }

  if (saveRowEditBtn) {
    saveRowEditBtn.addEventListener('click', function () {
      saveRowEdit()
    })
  }

  if (fileInput && fileName) {
    fileInput.addEventListener('change', function () {
      const file = fileInput.files && fileInput.files[0]
      fileName.innerText = file ? file.name : 'No file chosen'
      if (status) status.innerText = file ? 'Register selected. Ready to process.' : 'No register processed yet'
    })
  }

  if (processBtn) {
    processBtn.addEventListener('click', async function () {
      const file = fileInput.files && fileInput.files[0]

      if (!file) {
        if (status) status.innerText = 'Please upload a register first.'
        return
      }

      if (status) status.innerText = 'Sending register for review...'
      if (processingDots) processingDots.style.display = 'flex'

      try {
        const result = await requestRegisterReview(file)

        capturedRows = enrichRows(result.rows || [])
        window.capturedRows = capturedRows
        renderRows()

        if (status) status.innerText = result.message || 'Register processed. Please review highlighted fields before exporting.'
        if (summary) summary.innerText = capturedRows.length + ' rows loaded for review'
        if (processingDots) processingDots.style.display = 'none'
      } catch (error) {
        console.error('REGISTER REVIEW ERROR:', error)
        if (status) status.innerText = 'Register review failed. Please try again or contact admin.'
        if (processingDots) processingDots.style.display = 'none'
      }
    })
  }

  async function requestRegisterReview(file) {
    const fileBase64 = await fileToBase64(file)

    const response = await fetch('/api/register-review', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        fileBase64: fileBase64,
        mimeType: file.type || 'application/pdf',
        client: clientSelect ? clientSelect.value : ''
      })
    })

    if (!response.ok) {
      throw new Error('Register review API failed')
    }

    return response.json()
  }

  function fileToBase64(file) {
    return new Promise(function(resolve, reject) {
      const reader = new FileReader()

      reader.onload = function() {
        const result = String(reader.result || '')
        resolve(result.split(',')[1] || '')
      }

      reader.onerror = reject
      reader.readAsDataURL(file)
    })
  }

  window.renderRows = renderRows

  function renderRows() {
    if (!tableBody) return

    if (!capturedRows.length) {
      tableBody.innerHTML = '<tr><td colspan="8">No captured data yet</td></tr>'
      return
    }

    tableBody.innerHTML = capturedRows.map(function (row, index) {
      const risky = !row.name || !row.surname || !row.idNumber || row.idNumber.length < 10
      const rowClass = risky ? 'row-warning' : ''

      return `
        <tr class="${rowClass}" onclick="openRowModal(${index})" style="cursor:pointer;">
          <td>${index + 1}</td>
          <td contenteditable="true">${escapeHtml(row.name)}</td>
          <td contenteditable="true">${escapeHtml(row.surname)}</td>
          <td contenteditable="true">${escapeHtml(row.idNumber)}</td>
          <td contenteditable="true">${escapeHtml(row.contact)}</td>
          <td contenteditable="true">${escapeHtml(row.email)}</td>
          <td contenteditable="true">${escapeHtml(row.gender)}</td>
          <td><span class="status-pill status-duplicate">${escapeHtml(row.status || 'Review')}</span></td>
        </tr>
      `
    }).join('')
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;')
  }
})



let currentEditIndex = null;

function openRowModal(index) {
  const row = window.capturedRows[index];
  currentEditIndex = index;

  document.getElementById("editName").value = row.name || "";
  document.getElementById("editSurname").value = row.surname || "";
  document.getElementById("editIdNumber").value = row.idNumber || "";
  document.getElementById("editContact").value = row.contact || "";
  document.getElementById("editEmail").value = row.email || "N/A";
  document.getElementById("editGender").value = row.gender || "";
  document.getElementById("editRace").value = row.race || "B";
  document.getElementById("editEmploymentStatus").value = row.employmentStatus || "";
  document.getElementById("editIncomeRange").value = row.incomeRange || "";
  document.getElementById("editAge").value = row.age || calculateAge(row.idNumber || "");

  document.getElementById("rowModal").style.display = "flex";
}

function closeRowModal() {
  document.getElementById("rowModal").style.display = "none";
}

function saveRowEdit() {
  const row = window.capturedRows[currentEditIndex];

  row.name = document.getElementById("editName").value;
  row.surname = document.getElementById("editSurname").value;
  row.idNumber = document.getElementById("editIdNumber").value;
  row.contact = document.getElementById("editContact").value;
  row.email = document.getElementById("editEmail").value || "N/A";
  row.gender = document.getElementById("editGender").value;
  row.race = document.getElementById("editRace").value || "B";
  row.employmentStatus = document.getElementById("editEmploymentStatus").value;
  row.incomeRange = document.getElementById("editIncomeRange").value;
  row.age = document.getElementById("editAge").value || calculateAge(row.idNumber || "");
  row.status = rowNeedsReview(row) ? "Review" : "OK";

  capturedRows[currentEditIndex] = row;
  window.capturedRows = capturedRows;

  closeRowModal();

  window.renderRows();
}


function enrichRows(rows) {
  return rows.map(function(row) {
    row.email = row.email || "N/A";
    row.race = row.race || "B";
    row.age = row.age || calculateAge(row.idNumber || "");
    row.status = rowNeedsReview(row) ? "Review" : "OK";
    return row;
  });
}

function calculateAge(value) {
  const digits = String(value || "").replace(/\D/g, "");

  if (digits.length < 6) return "";

  const yy = Number(digits.slice(0, 2));
  const mm = Number(digits.slice(2, 4));
  const dd = Number(digits.slice(4, 6));

  if (!mm || !dd || mm > 12 || dd > 31) return "";

  const currentYear = new Date().getFullYear();
  const currentYY = Number(String(currentYear).slice(2));
  const century = yy <= currentYY ? 2000 : 1900;
  const birthYear = century + yy;

  let age = currentYear - birthYear;
  const today = new Date();
  const birthdayThisYear = new Date(currentYear, mm - 1, dd);

  if (today < birthdayThisYear) age -= 1;

  return age > 0 && age < 120 ? String(age) : "";
}

function rowNeedsReview(row) {
  const digits = String(row.idNumber || "").replace(/\D/g, "");
  return !row.name || !row.surname || digits.length < 10 || /\d/.test(row.name || "") || /\d/.test(row.surname || "");
}


function addRowBelow(index) {
  const newRow = {
    name: "",
    surname: "",
    idNumber: "",
    contact: "",
    email: "N/A",
    gender: "",
    race: "B",
    employmentStatus: "",
    incomeRange: "",
    age: "",
    status: "Review"
  };

  window.capturedRows.splice(index + 1, 0, newRow);
  capturedRows = window.capturedRows;
  window.renderRows();
}

function deleteRow(index) {
  const row = window.capturedRows[index] || {};
  const label = [row.name, row.surname].filter(Boolean).join(" ") || "this row";

  if (!confirm("Delete " + label + "?")) return;

  window.capturedRows.splice(index, 1);
  capturedRows = window.capturedRows;
  window.renderRows();
}
