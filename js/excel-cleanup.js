document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('excelFile')
  const fileNameText = document.getElementById('fileNameText')
  const processBtn = document.getElementById('processBtn')
  const exportBtn = document.getElementById('exportBtn')
  const sheetSelect = document.getElementById('sheetSelect')
  const ruleSelect = document.getElementById('ruleSelect')
  const ruleBadge = document.getElementById('ruleBadge')
  const duplicatesOnlyToggle = document.getElementById('duplicatesOnlyToggle')
  const hideRemovedToggle = document.getElementById('hideRemovedToggle')
  const statusText = document.getElementById('statusText')
  const summaryText = document.getElementById('summaryText')
  const table = document.getElementById('dataTable')
  const tableBody = document.getElementById('tableBody')

  let workbookRef = null
  let workbookMeta = { fileName: '', sheetNames: [] }
  let currentSheetName = ''
  let workingRows = []
  let activeRule = 'generic'

  fileInput.addEventListener('change', () => {
    const file = fileInput.files && fileInput.files[0]
    fileNameText.innerText = file ? file.name : 'No file chosen'
  })

  processBtn.addEventListener('click', () => {
    const file = fileInput.files && fileInput.files[0]

    if (!file) {
      statusText.innerText = 'Please choose a file first.'
      return
    }

    statusText.innerText = 'Reading workbook...'
    summaryText.innerText = 'Loading workbook...'
    tableBody.innerHTML = ''
    table.querySelector('thead').innerHTML = ''

    const reader = new FileReader()

    reader.onload = e => {
      const workbook = XLSX.read(e.target.result, { type: 'array' })
      workbookRef = workbook
      workbookMeta = {
        fileName: file.name,
        sheetNames: workbook.SheetNames.slice()
      }

      populateSheetSelector(workbook.SheetNames)

      const firstSheet = workbook.SheetNames[0]
      if (firstSheet) {
        sheetSelect.value = firstSheet
        loadSheet(firstSheet)
      }

      statusText.innerText = 'Workbook loaded.'
    }

    reader.onerror = () => {
      statusText.innerText = 'Could not read workbook.'
    }

    reader.readAsArrayBuffer(file)
  })

  sheetSelect.addEventListener('change', () => {
    loadSheet(sheetSelect.value)
  })

  ruleSelect.addEventListener('change', () => {
    if (currentSheetName) loadSheet(currentSheetName)
  })

  duplicatesOnlyToggle.addEventListener('change', () => {
    renderTable()
    updateSummary()
  })

  hideRemovedToggle.addEventListener('change', () => {
    renderTable()
    updateSummary()
  })

  exportBtn.addEventListener('click', () => {
    if (!workingRows.length) {
      statusText.innerText = 'No cleaned rows available to export yet.'
      return
    }

    const activeRows = workingRows.filter(row => row.rowState !== 'removed')
    const removedRows = workingRows.filter(row => row.rowState === 'removed')

    if (!activeRows.length) {
      statusText.innerText = 'No active rows available to export.'
      return
    }

    const headers = workingRows[0].headers || []

    const cleaned = activeRows.map(row => {
      const copy = {}
      headers.forEach(h => {
        copy[h] = row[h] || ''
      })
      return copy
    })

    const removed = removedRows.map(row => {
      const copy = {}
      headers.forEach(h => {
        copy[h] = row[h] || ''
      })
      copy['Cleanup Status'] = 'Removed duplicate'
      return copy
    })

    const summary = [
      { Metric: 'Workbook Sheet', Value: currentSheetName || 'Cleaned' },
      { Metric: 'Original Rows', Value: workingRows.length },
      { Metric: 'Cleaned Rows Exported', Value: activeRows.length },
      { Metric: 'Duplicate Rows Removed', Value: removedRows.length },
      { Metric: 'Rule Used', Value: activeRule === 'training_register' ? 'Training Register (Sessions)' : 'Generic workbook' },
      { Metric: 'Export Date', Value: new Date().toLocaleString() }
    ]

    const wb = XLSX.utils.book_new()

    const cleanedWs = XLSX.utils.json_to_sheet(cleaned)
    XLSX.utils.book_append_sheet(wb, cleanedWs, 'Cleaned Data')

    const summaryWs = XLSX.utils.json_to_sheet(summary)
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Cleanup Summary')

    if (removed.length) {
      const removedWs = XLSX.utils.json_to_sheet(removed)
      XLSX.utils.book_append_sheet(wb, removedWs, 'Removed Duplicates')
    }

    const fileName = `cleaned-${slugify(currentSheetName || 'sheet')}-${dateStamp()}.xlsx`
    XLSX.writeFile(wb, fileName)

    statusText.innerText = `Clean export ready: ${activeRows.length} rows exported, ${removedRows.length} duplicates removed.`
  })

  function populateSheetSelector(sheets) {
    sheetSelect.innerHTML = '<option value="">Select sheet</option>'
    sheets.forEach(name => {
      const option = document.createElement('option')
      option.value = name
      option.innerText = name
      sheetSelect.appendChild(option)
    })
  }

  function loadSheet(sheetName) {
    if (!workbookRef || !sheetName) return

    currentSheetName = sheetName

    const worksheet = workbookRef.Sheets[sheetName]
    const raw = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })

    if (!raw.length || raw.length < 3) {
      workingRows = []
      tableBody.innerHTML = ''
      table.querySelector('thead').innerHTML = ''
      summaryText.innerText = 'No usable rows found'
      updateRuleBadge('generic')
      return
    }

    const rawHeaderRow = raw[1] || []
    const headerIndexes = rawHeaderRow
      .map((h, i) => ({ h: String(h || '').trim(), i }))
      .filter(x => x.h && !x.h.includes('__EMPTY'))

    const headers = headerIndexes.map(x => x.h)
    const dataRows = raw.slice(2)

    const rows = dataRows
      .map((row, rowIndex) => {
        const obj = {
          _id: `${sheetName}-${rowIndex + 1}`,
          rowState: 'active',
          headers: headers.slice()
        }

        headerIndexes.forEach(({ h, i }) => {
          obj[h] = row[i] || ''
        })

        return obj
      })
      .filter(row => !isEmptyRow(row, headers))

    activeRule = resolveRule(headers, sheetName)
    updateRuleBadge(activeRule)

    workingRows = applyDuplicateFlags(rows, headers, activeRule)
    renderTable()
    updateSummary()
    statusText.innerText = `Sheet loaded: ${sheetName}. Duplicate rows were auto-marked for removal.`
  }

  function resolveRule(headers, sheetName) {
    if (ruleSelect.value && ruleSelect.value !== 'auto') {
      return ruleSelect.value
    }

    const normalizedHeaders = headers.map(h => normalize(h))
    const normalizedSheet = normalize(sheetName)
    const normalizedFile = normalize(workbookMeta.fileName)

    const hasSession =
      normalizedHeaders.some(h => h.includes('session')) ||
      normalizedHeaders.some(h => h.includes('ref number'))

    const hasFacilitator =
      normalizedHeaders.some(h => h.includes('facilitator'))

    const hasParticipant =
      normalizedHeaders.some(h => h === 'name') ||
      normalizedHeaders.some(h => h.includes('participant name')) ||
      normalizedHeaders.some(h => h.includes('full name'))

    const hasVenue =
      normalizedHeaders.some(h => h.includes('venue'))

    const looksLikeTrainingRegister =
      (hasSession && hasFacilitator && hasParticipant) ||
      normalizedSheet.includes('community') ||
      normalizedSheet.includes('worksite') ||
      normalizedSheet.includes('promotion') ||
      normalizedFile.includes('mimymo')

    if (looksLikeTrainingRegister && (hasVenue || hasSession || hasFacilitator)) {
      return 'training_register'
    }

    return 'generic'
  }

  function updateRuleBadge(rule) {
    const labelMap = {
      auto: 'Auto detect',
      training_register: 'Training Register (Sessions)',
      generic: 'Generic workbook'
    }
    ruleBadge.innerText = `Rule: ${labelMap[rule] || rule}`
  }

  function isEmptyRow(row, headers) {
    return headers.every(h => String(row[h] || '').trim() === '')
  }

  function applyDuplicateFlags(rows, headers, rule) {
    const map = new Map()

    rows.forEach(row => {
      const key = buildDuplicateKey(row, headers, rule)
      row._duplicateKey = key
      if (!map.has(key)) map.set(key, [])
      map.get(key).push(row)
    })

    rows.forEach(row => {
      const group = map.get(row._duplicateKey) || []
      const size = group.length
      const firstActive = group[0]

      row.isDuplicate = row._duplicateKey !== '__unique__' && size > 1
      row.groupSize = size
      row.keepChoice = row.isDuplicate && firstActive && firstActive._id === row._id

      if (row.isDuplicate && firstActive && firstActive._id !== row._id && row.rowState !== 'removed') {
        row.rowState = 'removed'
      }
    })

    return rows
  }

  function buildDuplicateKey(row, headers, rule) {
    const nameKey = findHeader(headers, ['Name', 'Full Name', 'Participant Name'])
    const sessionKey = findHeader(headers, ['Session Code/ Ref Number', 'Session Code', 'Ref Number', 'Session Code / Ref Number'])
    const idKey = findHeader(headers, ['ID Number', 'ID No', 'ID', 'Identity Number'])
    const emailKey = findHeader(headers, ['Email', 'E-mail', 'Email Address', 'Mail'])
    const dateKey = findHeader(headers, ['Date', 'Session Date'])
    const phoneKey = findHeader(headers, ['Phone', 'Cell', 'Cell Number', 'Mobile', 'Mobile Number'])

    const name = normalize(row[nameKey])
    const session = normalize(row[sessionKey])
    const id = normalizeDigits(row[idKey])
    const email = normalize(row[emailKey])
    const date = normalize(row[dateKey])
    const phone = normalizeDigits(row[phoneKey])

    if (rule === 'training_register') {
      if (session && name) return `session-name|${session}|${name}`
      if (date && name) return `date-name|${date}|${name}`
      if (id) return `id|${id}`
      if (email) return `email|${email}`
      if (phone) return `phone|${phone}`
      if (name) return `name|${name}`
      return '__unique__'
    }

    if (rule === 'generic') {
      if (id) return `id|${id}`
      if (email) return `email|${email}`
      if (session && name) return `session-name|${session}|${name}`
      if (name) return `name|${name}`
      return '__unique__'
    }

    return '__unique__'
  }

  function findHeader(headers, candidates) {
    const normalizedHeaders = headers.map(h => ({ raw: h, norm: normalize(h) }))
    for (const candidate of candidates) {
      const target = normalize(candidate)
      const found = normalizedHeaders.find(h => h.norm === target)
      if (found) return found.raw
    }
    return ''
  }

  function normalize(value) {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ')
  }

  function normalizeDigits(value) {
    return String(value || '').replace(/\D+/g, '')
  }

  function getVisibleRows() {
    let rows = workingRows.slice()

    if (duplicatesOnlyToggle.checked) {
      rows = rows.filter(r => r.isDuplicate)
    }

    if (hideRemovedToggle.checked) {
      rows = rows.filter(r => r.rowState !== 'removed')
    }

    return rows
  }

  function renderTable() {
    tableBody.innerHTML = ''

    if (!workingRows.length) {
      table.querySelector('thead').innerHTML = ''
      return
    }

    const headers = workingRows[0].headers
    const visibleRows = getVisibleRows()

    table.querySelector('thead').innerHTML = `
      <tr>
        <th>#</th>
        ${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
        <th class="status-col">Status</th>
        <th class="actions-col">Actions</th>
      </tr>
    `

    visibleRows.forEach((row, index) => {
      const tr = document.createElement('tr')
      tr.dataset.id = row._id

      if (row.rowState === 'removed') {
        tr.classList.add('row-removed')
      } else if (row.isDuplicate) {
        tr.classList.add('row-duplicate')
      }

      tr.innerHTML = `
        <td class="index-col">${index + 1}</td>
        ${headers.map(h => `
          <td class="editable-cell" contenteditable="${row.rowState === 'removed' ? 'false' : 'true'}">${escapeHtml(row[h])}</td>
        `).join('')}
        <td>${getStatusHtml(row)}</td>
        <td>${getActionsHtml(row)}</td>
      `

      tableBody.appendChild(tr)
    })

    bindRowActions()
    bindEditableCells(headers)
  }

  function getStatusHtml(row) {
    if (row.rowState === 'removed') {
      return `<span class="status-pill status-removed">Removed</span>`
    }

    if (row.keepChoice) {
      return `<span class="status-pill status-kept">Kept</span>`
    }

    if (row.isDuplicate) {
      return `<span class="status-pill status-duplicate">Duplicate (${row.groupSize})</span>`
    }

    return `<span class="status-pill status-unique">Unique</span>`
  }

  function getActionsHtml(row) {
    if (row.rowState === 'removed') {
      return `
        <div class="row-actions">
          <button class="row-btn keep" data-action="restore" data-id="${row._id}">Restore</button>
        </div>
      `
    }

    if (row.isDuplicate) {
      return `
        <div class="row-actions">
          <button class="row-btn keep" data-action="keep" data-id="${row._id}">Keep</button>
          <button class="row-btn remove" data-action="remove" data-id="${row._id}">Remove</button>
        </div>
      `
    }

    return `
      <div class="row-actions">
        <button class="row-btn remove" data-action="remove" data-id="${row._id}">Remove</button>
      </div>
    `
  }

  function bindRowActions() {
    tableBody.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', () => {
        const action = btn.dataset.action
        const id = btn.dataset.id
        const row = workingRows.find(r => r._id === id)
        if (!row) return

        if (action === 'remove') {
          row.rowState = 'removed'
          row.keepChoice = false
        }

        if (action === 'restore') {
          row.rowState = 'active'
        }

        if (action === 'keep') {
          workingRows.forEach(r => {
            if (r._duplicateKey === row._duplicateKey) {
              r.keepChoice = false
            }
          })
          row.keepChoice = true
          row.rowState = 'active'
        }

        renderTable()
        updateSummary()
      })
    })
  }

  function bindEditableCells(headers) {
    tableBody.querySelectorAll('tr').forEach(tr => {
      const row = workingRows.find(r => r._id === tr.dataset.id)
      if (!row || row.rowState === 'removed') return

      const tds = tr.querySelectorAll('td')
      headers.forEach((h, i) => {
        const td = tds[i + 1]
        if (!td) return

        td.addEventListener('blur', () => {
          row[h] = td.innerText.trim()
          workingRows = applyDuplicateFlags(workingRows, headers, activeRule)
          renderTable()
          updateSummary()
        })
      })
    })
  }

  function updateSummary() {
    const total = workingRows.length
    const active = workingRows.filter(r => r.rowState !== 'removed').length
    const removed = workingRows.filter(r => r.rowState === 'removed').length
    const duplicates = workingRows.filter(r => r.isDuplicate && r.rowState !== 'removed').length
    const kept = workingRows.filter(r => r.keepChoice && r.rowState !== 'removed').length
    const visible = getVisibleRows().length

    summaryText.innerText = `${total} rows • ${active} active • ${removed} removed • ${duplicates} duplicate rows • ${kept} kept • ${visible} visible`
  }

  function dateStamp() {
    const d = new Date()
    const y = d.getFullYear()
    const m = String(d.getMonth() + 1).padStart(2, '0')
    const day = String(d.getDate()).padStart(2, '0')
    const h = String(d.getHours()).padStart(2, '0')
    const min = String(d.getMinutes()).padStart(2, '0')
    return `${y}${m}${day}-${h}${min}`
  }

  function slugify(value) {
    return String(value || 'sheet')
      .toLowerCase()
      .trim()
      .replace(/\s+/g, '-')
      .replace(/[^a-z0-9-_]/g, '')
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
  }
})
