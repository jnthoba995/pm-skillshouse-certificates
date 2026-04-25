document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('registerFile')
  const fileName = document.getElementById('registerFileName')
  const status = document.getElementById('registerStatus')
  const processBtn = document.getElementById('processRegisterBtn')
  const summary = document.getElementById('summaryText')
  const tableBody = document.getElementById('tableBody')

  let capturedRows = []

  if (fileInput && fileName) {
    fileInput.addEventListener('change', function () {
      const file = fileInput.files && fileInput.files[0]
      fileName.innerText = file ? file.name : 'No file chosen'
      if (status) status.innerText = file ? 'Register selected. Ready to process.' : 'No register processed yet'
    })
  }

  if (processBtn) {
    processBtn.addEventListener('click', function () {
      capturedRows = [
        {
          name: 'John',
          surname: 'Doe',
          idNumber: '',
          contact: '',
          email: '',
          gender: '',
          status: 'Needs review'
        },
        {
          name: '',
          surname: 'Smith',
          idNumber: '123',
          contact: '',
          email: '',
          gender: '',
          status: 'Needs review'
        }
      ]

      renderRows()
      if (status) status.innerText = 'Register processed. Please review highlighted fields before exporting.'
      if (summary) summary.innerText = capturedRows.length + ' rows loaded for review'
    })
  }

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
        <tr class="${rowClass}">
          <td>${index + 1}</td>
          <td contenteditable="true">${escapeHtml(row.name)}</td>
          <td contenteditable="true">${escapeHtml(row.surname)}</td>
          <td contenteditable="true">${escapeHtml(row.idNumber)}</td>
          <td contenteditable="true">${escapeHtml(row.contact)}</td>
          <td contenteditable="true">${escapeHtml(row.email)}</td>
          <td contenteditable="true">${escapeHtml(row.gender)}</td>
          <td><span class="status-pill status-duplicate">${escapeHtml(row.status)}</span></td>
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
