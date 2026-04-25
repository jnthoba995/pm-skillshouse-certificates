document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('registerFile')
  const fileName = document.getElementById('registerFileName')
  const status = document.getElementById('registerStatus')
  const processBtn = document.getElementById('processRegisterBtn')
  const summary = document.getElementById('summaryText')
  const tableBody = document.getElementById('tableBody')
  const clientSelect = document.getElementById('registerClient')

  let capturedRows = []

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

      try {
        const result = await requestRegisterReview(file)

        capturedRows = result.rows || []
        renderRows()

        if (status) status.innerText = result.message || 'Register processed. Please review highlighted fields before exporting.'
        if (summary) summary.innerText = capturedRows.length + ' rows loaded for review'
      } catch (error) {
        console.error('REGISTER REVIEW ERROR:', error)
        if (status) status.innerText = 'Register review failed. Please try again or contact admin.'
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
