document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('excelFile')
  const fileNameText = document.getElementById('fileNameText')
  const processBtn = document.getElementById('processBtn')
  const exportBtn = document.getElementById('exportBtn')
  const saveDraftBtn = document.getElementById('saveDraftBtn')
  const restoreDraftBtn = document.getElementById('restoreDraftBtn')
  const sheetSelect = document.getElementById('sheetSelect')
  const ruleSelect = document.getElementById('ruleSelect')
  const ruleBadge = document.getElementById('ruleBadge')
  const duplicatesOnlyToggle = document.getElementById('duplicatesOnlyToggle')
  const hideRemovedToggle = document.getElementById('hideRemovedToggle')
  const riskOnlyToggle = document.getElementById('riskOnlyToggle')
  const statusText = document.getElementById('statusText')
  const summaryText = document.getElementById('summaryText')
  const table = document.getElementById('dataTable')
  const tableBody = document.getElementById('tableBody')

  let workbookRef = null
  let workbookMeta = { fileName: '', sheetNames: [] }
  let currentSheetName = ''
  let workingRows = []
  let allWorkbookRows = []
  let activeSheetFilter = 'All'
  let activeRule = 'generic'
  let duplicateViewMode = 'auto'
  let duplicateInvestigationBaseRows = []
  let duplicateInvestigationFilters = { id: false, cellphone: false, province: false, name: false }


  
  

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
        sheetSelect.value = '__ALL_SHEETS__'
        loadSheet('__ALL_SHEETS__')
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


  saveDraftBtn.addEventListener('click', () => {
    if (!workingRows.length) {
      statusText.innerText = 'No cleanup session available to save yet.'
      return
    }

    const payload = {
      savedAt: new Date().toISOString(),
      currentSheetName,
      activeRule,
      workingRows: allWorkbookRows.length ? allWorkbookRows : workingRows
    }

    localStorage.setItem('pm_cleanup_draft', JSON.stringify(payload))
    statusText.innerText = `Draft saved at ${new Date(payload.savedAt).toLocaleString()}.`
  })

  restoreDraftBtn.addEventListener('click', () => {
    const saved = localStorage.getItem('pm_cleanup_draft')

    if (!saved) {
      statusText.innerText = 'No saved cleanup draft found.'
      return
    }

    const payload = JSON.parse(saved)

    currentSheetName = payload.currentSheetName || ''
    activeRule = payload.activeRule || 'generic'
    allWorkbookRows = payload.workingRows || []
    workingRows = allWorkbookRows.slice()
    activeSheetFilter = 'All'
    renderCleanupSheetFilters()
    renderDuplicateInvestigationPanel()

    updateRuleBadge(activeRule)
    renderTable()
    updateSummary()
    updateAutoFixSuggestions()

    statusText.innerText = `Draft restored from ${new Date(payload.savedAt).toLocaleString()}.`
  })


  ruleSelect.addEventListener('change', () => {
    if (currentSheetName) loadSheet(currentSheetName)
  })

  duplicatesOnlyToggle.addEventListener('change', () => {
    renderTable()
    updateSummary()
    updateAutoFixSuggestions()
  })

  hideRemovedToggle.addEventListener('change', () => {
    renderTable()
  })

  if (riskOnlyToggle) {
    riskOnlyToggle.addEventListener('change', () => {
      renderTable()
    })
  }

  // continue
  hideRemovedToggle.addEventListener('change', () => {
    renderTable()
    updateSummary()
    updateAutoFixSuggestions()
  })

  exportBtn.addEventListener('click', () => {
    if (!workingRows.length) {
      statusText.innerText = 'No cleaned rows available to export yet.'
      return
    }

    const exportSourceRows = allWorkbookRows.length ? allWorkbookRows : workingRows
    const activeRows = exportSourceRows.filter(row => row.rowState !== 'removed')
    const removedRows = exportSourceRows.filter(row => row.rowState === 'removed')

    if (!activeRows.length) {
      statusText.innerText = 'No active rows available to export.'
      return
    }

    const wb = XLSX.utils.book_new()

    const sourceSheets = Array.from(new Set(exportSourceRows.map(row => row['Source Sheet']).filter(Boolean)))

    const sheetsToExport = sourceSheets.length ? sourceSheets : [currentSheetName || 'Cleaned Data']

    sheetsToExport.forEach(sheetName => {
      const rowsForSheet = activeRows.filter(row => {
        if (!sourceSheets.length) return true
        return row['Source Sheet'] === sheetName
      })

      if (!rowsForSheet.length) return

      const sheetHeaders = getHeadersForSelectedTab(rowsForSheet).filter(h => h !== 'Actions')

      const cleaned = rowsForSheet.map(row => {
        const copy = {}
        sheetHeaders.forEach(h => {
          if (h === 'Source Sheet' && sourceSheets.length === 1) return
          copy[h] = row[h] || ''
        })
        return copy
      })

      const safeSheetName = String(sheetName || 'Cleaned Data').replace(/[\\\/\?\*\[\]\:]/g, ' ').slice(0, 31) || 'Cleaned Data'
      const cleanedWs = XLSX.utils.json_to_sheet(cleaned)
      XLSX.utils.book_append_sheet(wb, cleanedWs, safeSheetName)
    })

    const removedHeaders = getHeadersForAllTabs(removedRows.length ? removedRows : exportSourceRows).filter(h => h !== 'Actions')

    const removed = removedRows.map(row => {
      const copy = {}
      removedHeaders.forEach(h => {
        copy[h] = row[h] || ''
      })
      copy['Cleanup Status'] = 'Marked for review'
      copy['Duplicate Reason'] = row.duplicateReason || row.similarityReason || ''
      return copy
    })

    const summary = [
      { Metric: 'Workbook Sheet', Value: currentSheetName || 'Cleaned' },
      { Metric: 'Tabs Exported', Value: sheetsToExport.length },
      { Metric: 'Original Rows', Value: exportSourceRows.length },
      { Metric: 'Cleaned Rows Exported', Value: activeRows.length },
      { Metric: 'Rows Marked for Review', Value: removedRows.length },
      { Metric: 'Rule Used', Value: activeRule === 'training_register' ? 'Training Register (Sessions)' : 'Generic workbook' },
      { Metric: 'Duplicate View Used', Value: duplicateViewMode },
      { Metric: 'Export Date', Value: new Date().toLocaleString() }
    ]

    const summaryWs = XLSX.utils.json_to_sheet(summary)
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Cleanup Summary')

    if (removed.length) {
      const removedWs = XLSX.utils.json_to_sheet(removed)
      XLSX.utils.book_append_sheet(wb, removedWs, 'Review List')
    }

    const fileName = `cleaned-${slugify(currentSheetName || 'workbook')}-${dateStamp()}.xlsx`
    XLSX.writeFile(wb, fileName)

    statusText.innerText = `Clean export ready: ${activeRows.length} rows exported across ${sheetsToExport.length} tab(s).`
  })

  function populateSheetSelector(sheets) {
    sheetSelect.innerHTML = '<option value="">Select sheet</option>'

    const allOption = document.createElement('option')
    allOption.value = '__ALL_SHEETS__'
    allOption.innerText = 'All Workbook Tabs'
    sheetSelect.appendChild(allOption)

    sheets.forEach(name => {
      const option = document.createElement('option')
      option.value = name
      option.innerText = name
      sheetSelect.appendChild(option)
    })
  }

  function detectHeaderRow(raw) {
    const signals = [
      'date', 'facilitator', 'session code', 'ref number', 'venue',
      'province', 'municipality', 'urban', 'rural', 'gender', 'sex',
      'age', 'employment', 'employed', 'name', 'surname', 'id number',
      'cell', 'phone', 'email', 'town', 'city', 'ward', 'company'
    ]

    let bestIndex = 0
    let bestScore = -1

    raw.slice(0, 30).forEach((row, i) => {
      const cells = (row || []).map(v => String(v || '').trim()).filter(Boolean)
      const joined = cells.join(' ').toLowerCase()

      if (!cells.length) return

      let score = cells.length

      signals.forEach(sig => {
        if (joined.includes(sig)) score += 15
      })

      if (joined.includes('residential details')) score -= 20
      if (joined.includes('personal details')) score -= 20
      if (joined.includes('attendance register')) score -= 20

      if (score > bestScore) {
        bestScore = score
        bestIndex = i
      }
    })

    return bestIndex
  }


  function isDateHeader(header) {
    const h = String(header || '').toLowerCase()
    return h === 'date' || h.includes('session date') || h.includes('training date')
  }

  function formatExcelDateValue(value, header) {
    if (!isDateHeader(header)) return value

    if (value instanceof Date && !isNaN(value)) {
      return value.toISOString().slice(0, 10)
    }

    const raw = String(value || '').trim()
    if (!raw) return ''

    if (/^\d{5}$/.test(raw)) {
      const serial = Number(raw)
      const utcDays = Math.floor(serial - 25569)
      const utcValue = utcDays * 86400
      const dateInfo = new Date(utcValue * 1000)
      return dateInfo.toISOString().slice(0, 10)
    }

    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(raw)) {
      const parts = raw.split('/')
      const month = parts[0].padStart(2, '0')
      const day = parts[1].padStart(2, '0')
      const year = parts[2].length === 2 ? '20' + parts[2] : parts[2]
      return `${year}-${month}-${day}`
    }

    return value
  }


  function rowsFromWorksheet(sheetName) {
    const worksheet = workbookRef.Sheets[sheetName]
    const raw = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', blankrows: false })

    if (!raw.length || raw.length < 2) {
      return { headers: [], rows: [] }
    }

    const headerRowIndex = detectHeaderRow(raw)
    const rawHeaderRow = raw[headerRowIndex] || []

    const headerIndexes = rawHeaderRow
      .map((h, i) => ({ h: String(h || '').trim(), i }))
      .filter(x => x.h && !x.h.includes('__EMPTY'))

    const dataHeaders = headerIndexes.map(x => x.h)

    if (!dataHeaders.length) {
      return { headers: [], rows: [] }
    }

    const knownSignals = [
      'date', 'facilitator', 'session', 'ref', 'venue', 'province',
      'municipality', 'urban', 'rural', 'gender', 'employment',
      'age', 'name', 'surname', 'id', 'phone', 'cell', 'email'
    ]

    const headerText = dataHeaders.join(' ').toLowerCase()
    const signalCount = knownSignals.filter(sig => headerText.includes(sig)).length

    if (signalCount < 2) {
      return { headers: [], rows: [] }
    }

    const headers = ['Source Sheet'].concat(dataHeaders.filter(h => h !== 'Source Sheet'))
    const dataRows = raw.slice(headerRowIndex + 1)

    const rows = dataRows
      .map((row, rowIndex) => {
        const realExcelRow = headerRowIndex + rowIndex + 2
        const obj = {
          _id: `${sheetName}-${rowIndex + 1}`,
          rowState: 'active',
          headers: headers.slice(),
          sheetHeaders: headers.slice(),
          'Real Excel Row': realExcelRow,
          'Original Excel Row': realExcelRow,
          'Source Sheet': sheetName
        }

        headerIndexes.forEach(({ h, i }) => {
          obj[h] = formatExcelDateValue(row[i] || '', h)
        })

        return obj
      })
      .filter(row => {
        const values = dataHeaders
          .map(h => String(row[h] || '').trim())
          .filter(Boolean)

        if (values.length < 2) return false

        const joined = values.join(' ').toLowerCase()
        if (joined.includes('residential details')) return false
        if (joined.includes('personal details')) return false
        if (joined.includes('attendance register')) return false

        return true
      })

    return { headers, rows }
  }

  function renderCleanupSheetFilters() {
    const reviewHead = document.querySelector('.review-head')
    if (!reviewHead) return

    let wrap = document.getElementById('cleanupSheetFilterWrap')

    if (!wrap) {
      reviewHead.insertAdjacentHTML('beforeend', `
        <div id="cleanupSheetFilterWrap" class="cleanup-sheet-filter-wrap">
        </div>
      `)
      wrap = document.getElementById('cleanupSheetFilterWrap')
    }

    const sourceRows = allWorkbookRows.length ? allWorkbookRows : workingRows
    const sheets = Array.from(new Set(sourceRows.map(row => row['Source Sheet']).filter(Boolean)))

    if (!sheets.length) {
      wrap.innerHTML = ''
      return
    }

    wrap.innerHTML =
      '<span class="cleanup-filter-label">View:</span>' +
      '<button type="button" class="cleanup-sheet-filter active" data-sheet="All">All tabs</button>' +
      sheets.map(name => `<button type="button" class="cleanup-sheet-filter" data-sheet="${escapeHtml(name)}">${escapeHtml(name)}</button>`).join('')

    wrap.querySelectorAll('.cleanup-sheet-filter').forEach(btn => {
      btn.addEventListener('click', () => {
        activeSheetFilter = btn.getAttribute('data-sheet') || 'All'

        wrap.querySelectorAll('.cleanup-sheet-filter').forEach(b => b.classList.remove('active'))
        btn.classList.add('active')

        applyCleanupSheetFilter()
      })
    })
  }



  function renderDuplicateInvestigationPanel() {
    const reviewHead = document.querySelector('.review-head')
    if (!reviewHead) return

    let panel = document.getElementById('duplicateInvestigationPanel')

    if (!panel) {
      reviewHead.insertAdjacentHTML('afterend', `
        <div id="duplicateInvestigationPanel" class="duplicate-investigation-panel">
          <div class="duplicate-investigation-head">
            <div>
              <strong>Duplicate Investigation</strong>
              <span>Choose one or more checks. Filters can be used in any order.</span>
            </div>
            <button type="button" id="duplicateResetBtn" class="duplicate-filter-btn reset">Reset</button>
          </div>

          <div class="duplicate-filter-row">
            <button type="button" class="duplicate-filter-btn id" data-dup-filter="id">ID Number</button>
            <button type="button" class="duplicate-filter-btn cell" data-dup-filter="cellphone">Cellphone</button>
            <button type="button" class="duplicate-filter-btn province" data-dup-filter="province">Province</button>
            <button type="button" class="duplicate-filter-btn name" data-dup-filter="name">Name & Surname</button>
          </div>

          <div id="duplicateInvestigationStatus" class="duplicate-investigation-status">
            Upload Excel, then select a check to narrow possible duplicate records.
          </div>
        </div>
      `)

      panel = document.getElementById('duplicateInvestigationPanel')
    }

    panel.querySelectorAll('[data-dup-filter]').forEach(btn => {
      btn.onclick = function () {
        const key = btn.getAttribute('data-dup-filter')

        if (key !== 'id' && key !== 'cellphone' && key !== 'province' && key !== 'name') {
          btn.classList.toggle('active')
          const status = document.getElementById('duplicateInvestigationStatus')
          if (status) {
            status.textContent = 'ID Number and Cellphone checks are active. Province and Name will follow in the next patches.'
          }
          btn.classList.remove('active')
          return
        }

        btn.classList.toggle('active')
        duplicateInvestigationFilters[key] = btn.classList.contains('active')
        applyDuplicateInvestigationFilters()
      }
    })

    const resetBtn = document.getElementById('duplicateResetBtn')
    if (resetBtn) {
      resetBtn.onclick = function () {
        panel.querySelectorAll('[data-dup-filter]').forEach(btn => btn.classList.remove('active'))
        duplicateInvestigationFilters = { id: false, cellphone: false, province: false, name: false }
        resetDuplicateInvestigationFilters()
      }
    }
  }


  function clearDuplicateInvestigationFlags(rows) {
    rows.forEach(row => {
      row.duplicateInvestigationReasons = []
      row.duplicateInvestigationLevels = []
      row.duplicateInvestigationTags = []
      row.duplicateInvestigationActive = false
      row.duplicateInvestigationGroup = ''
      row.duplicateInvestigationGroupReason = ''
    })
  }

  function getInvestigationHeaders(rows) {
    if (rows && rows[0] && rows[0].headers) return rows[0].headers
    if (rows && rows[0] && rows[0].__headers) return rows[0].__headers
    if (rows && rows.length) {
      return Object.keys(rows[0]).filter(key => !String(key).startsWith('__'))
    }
    return []
  }

  function getInvestigationIdValue(row, headers) {
    if (typeof getIdentityFieldValue === 'function') {
      return getIdentityFieldValue(row, headers, 'id')
    }

    if (typeof valueForHeader === 'function') {
      return valueForHeader(row, headers, [
        'ID Number', 'ID No', 'ID', 'Identity Number', 'ID/Passport Number',
        'ID No./Date of Birth', 'ID Number / Date of Birth', 'ID Number / DOB',
        'Learner ID', 'Learner ID Number'
      ])
    }

    const aliases = [
      'idnumber', 'idno', 'id', 'identitynumber', 'identityno',
      'idnumberdob', 'idnumberdateofbirth', 'learnerid', 'learneridnumber'
    ]

    const keys = Object.keys(row || {})

    for (const key of keys) {
      const normal = String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '')
      if (aliases.includes(normal)) return row[key]
    }

    return ''
  }

  function markInvestigationRow(row, tag, level, reason) {
    row.duplicateInvestigationActive = true

    if (!row.duplicateInvestigationTags) row.duplicateInvestigationTags = []
    if (!row.duplicateInvestigationLevels) row.duplicateInvestigationLevels = []
    if (!row.duplicateInvestigationReasons) row.duplicateInvestigationReasons = []

    if (!row.duplicateInvestigationTags.includes(tag)) row.duplicateInvestigationTags.push(tag)
    if (!row.duplicateInvestigationLevels.includes(level)) row.duplicateInvestigationLevels.push(level)
    if (!row.duplicateInvestigationReasons.includes(reason)) row.duplicateInvestigationReasons.push(reason)
  }

  function findIdInvestigationRows(rows) {
    const headers = getInvestigationHeaders(rows)
    const exactGroups = {}
    const nearGroups = {}
    const matchGroups = []
    let matchGroupCounter = 1

    rows.forEach(row => {
      const raw = getInvestigationIdValue(row, headers)
      const digits = normalizeDigits(raw)

      if (!digits || digits.length < 6) return

      const exactKey = digits
      if (!exactGroups[exactKey]) exactGroups[exactKey] = []
      exactGroups[exactKey].push(row)

      if (digits.length >= 10) {
        const nearKey = digits.slice(0, Math.max(6, digits.length - 4))
        if (!nearGroups[nearKey]) nearGroups[nearKey] = []
        nearGroups[nearKey].push(row)
      }
    })

    const matched = new Set()

    Object.keys(exactGroups).forEach(key => {
      const group = exactGroups[key]
      if (group.length < 2) return

      const groupName = 'ID Group ' + matchGroupCounter++
      matchGroups.push({ groupName, rows: group, reason: 'Exact ID match' })

      group.forEach(row => {
        matched.add(row)
        row.duplicateInvestigationGroup = groupName
        row.duplicateInvestigationGroupReason = 'Exact ID match'
        markInvestigationRow(row, 'id', 'exact', 'ID Number exact match')
      })
    })

    Object.keys(nearGroups).forEach(key => {
      const group = nearGroups[key]
      if (group.length < 2) return

      const uniqueIds = new Set(group.map(row => normalizeDigits(getInvestigationIdValue(row, getInvestigationHeaders(group)))))
      if (uniqueIds.size < 2) return

      const ungroupedRows = group.filter(row => !row.duplicateInvestigationGroup)
      if (ungroupedRows.length < 2) return

      const groupName = 'ID Group ' + matchGroupCounter++
      matchGroups.push({ groupName, rows: ungroupedRows, reason: 'Near ID match' })

      ungroupedRows.forEach(row => {
        matched.add(row)
        row.duplicateInvestigationGroup = groupName
        row.duplicateInvestigationGroupReason = 'Near ID match - last digits may differ'
        markInvestigationRow(row, 'id', 'near', 'Possible ID issue - last 1 to 4 digits may differ')
      })
    })

    return rows
      .filter(row => matched.has(row))
      .sort((a, b) => {
        const ag = a.duplicateInvestigationGroup || ''
        const bg = b.duplicateInvestigationGroup || ''
        if (ag !== bg) return ag.localeCompare(bg, undefined, { numeric: true })
        const aid = normalizeDigits(getInvestigationIdValue(a, headers))
        const bid = normalizeDigits(getInvestigationIdValue(b, headers))
        return aid.localeCompare(bid)
      })
  }


  function getInvestigationCellValue(row, headers) {
    if (typeof getIdentityFieldValue === 'function') {
      return getIdentityFieldValue(row, headers, 'cell')
    }

    if (typeof valueForHeader === 'function') {
      return valueForHeader(row, headers, [
        'Cellphone', 'Cellphone Number', 'Cell Phone', 'Cell Number',
        'Cell', 'Phone', 'Phone Number', 'Contact Number', 'Mobile', 'Mobile Number',
        'Learner Cellphone', 'Learner Cellphone Number', 'Learner Contact Number'
      ])
    }

    const aliases = [
      'cellphone', 'cellphonenumber', 'cellphonecellnumber', 'cellnumber',
      'cell', 'phone', 'phonenumber', 'contactnumber', 'mobile', 'mobilenumber',
      'learnercellphone', 'learnercellphonenumber', 'learnercontactnumber'
    ]

    const keys = Object.keys(row || {})

    for (const key of keys) {
      const normal = String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '')
      if (aliases.includes(normal)) return row[key]
    }

    return ''
  }

  function findCellphoneInvestigationRows(rows) {
    const headers = getInvestigationHeaders(rows)
    const exactGroups = {}
    const nearGroups = {}
    const matched = new Set()
    let groupCounter = 1

    rows.forEach(row => {
      const raw = getInvestigationCellValue(row, headers)
      let digits = normalizeDigits(raw)

      if (!digits || digits.length < 6) return

      if (digits.length > 10 && digits.startsWith('27')) {
        digits = '0' + digits.slice(2)
      }

      const exactKey = digits
      if (!exactGroups[exactKey]) exactGroups[exactKey] = []
      exactGroups[exactKey].push(row)

      if (digits.length >= 7) {
        const nearKey = digits.slice(0, Math.max(5, digits.length - 4))
        if (!nearGroups[nearKey]) nearGroups[nearKey] = []
        nearGroups[nearKey].push(row)
      }
    })

    Object.keys(exactGroups).forEach(key => {
      const group = exactGroups[key]
      if (group.length < 2) return

      const groupName = 'Cell Group ' + groupCounter++

      group.forEach(row => {
        matched.add(row)
        if (!row.duplicateInvestigationGroup) row.duplicateInvestigationGroup = groupName
        if (!row.duplicateInvestigationGroupReason) row.duplicateInvestigationGroupReason = 'Exact cellphone match'
        markInvestigationRow(row, 'cellphone', 'exact', 'Cellphone exact match')
      })
    })

    Object.keys(nearGroups).forEach(key => {
      const group = nearGroups[key]
      if (group.length < 2) return

      const uniqueCells = new Set(group.map(row => normalizeDigits(getInvestigationCellValue(row, headers))))
      if (uniqueCells.size < 2) return

      const groupName = 'Cell Group ' + groupCounter++

      group.forEach(row => {
        matched.add(row)
        if (!row.duplicateInvestigationGroup) row.duplicateInvestigationGroup = groupName
        if (!row.duplicateInvestigationGroupReason) row.duplicateInvestigationGroupReason = 'Near cellphone match - last digits may differ'
        markInvestigationRow(row, 'cellphone', 'near', 'Possible cellphone issue - last 1 to 4 digits may differ')
      })
    })

    return rows
      .filter(row => matched.has(row))
      .sort((a, b) => {
        const ag = a.duplicateInvestigationGroup || ''
        const bg = b.duplicateInvestigationGroup || ''
        if (ag !== bg) return ag.localeCompare(bg, undefined, { numeric: true })

        const ac = normalizeDigits(getInvestigationCellValue(a, headers))
        const bc = normalizeDigits(getInvestigationCellValue(b, headers))
        return ac.localeCompare(bc)
      })
  }



  function getInvestigationProvinceValue(row, headers) {
    if (typeof valueForHeader === 'function') {
      return valueForHeader(row, headers, [
        'Province', 'Home Province', 'Residential Province', 'Province Name',
        'Area Province', 'Region', 'Home Region'
      ])
    }

    const aliases = [
      'province', 'homeprovince', 'residentialprovince', 'provincename',
      'areaprovince', 'region', 'homeregion'
    ]

    const keys = Object.keys(row || {})

    for (const key of keys) {
      const normal = String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '')
      if (aliases.includes(normal)) return row[key]
    }

    return ''
  }

  function normalizeProvinceValue(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .replace(/[^a-z ]/g, '')
      .trim()
  }

  function findProvinceInvestigationRows(rows) {
    const headers = getInvestigationHeaders(rows)
    const groups = {}
    const matched = new Set()
    let groupCounter = 1

    rows.forEach(row => {
      const province = normalizeProvinceValue(getInvestigationProvinceValue(row, headers))
      if (!province || province.length < 3) return

      if (!groups[province]) groups[province] = []
      groups[province].push(row)
    })

    Object.keys(groups).forEach(key => {
      const group = groups[key]
      if (group.length < 2) return

      const groupName = 'Province Group ' + groupCounter++

      group.forEach(row => {
        matched.add(row)
        if (!row.duplicateInvestigationGroup) row.duplicateInvestigationGroup = groupName
        if (!row.duplicateInvestigationGroupReason) row.duplicateInvestigationGroupReason = 'Same province'
        markInvestigationRow(row, 'province', 'exact', 'Same province')
      })
    })

    return rows
      .filter(row => matched.has(row))
      .sort((a, b) => {
        const ag = a.duplicateInvestigationGroup || ''
        const bg = b.duplicateInvestigationGroup || ''
        if (ag !== bg) return ag.localeCompare(bg, undefined, { numeric: true })

        const ap = normalizeProvinceValue(getInvestigationProvinceValue(a, headers))
        const bp = normalizeProvinceValue(getInvestigationProvinceValue(b, headers))
        return ap.localeCompare(bp)
      })
  }



  function normalizeInvestigationName(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/[^a-z ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
  }

  function getInvestigationFullNameValue(row, headers) {
    const combined = getRowNameSignal(row, headers)
    if (combined) return combined

    const full = valueForHeader(row, headers, [
      'Full Name', 'FullName', 'Participant Name', 'Learner Full Name'
    ])

    if (full) return full

    const first = valueForHeader(row, headers, [
      'Name', 'First Name', 'FirstName', 'Learner Name', 'Learner First Name'
    ])

    const last = valueForHeader(row, headers, [
      'Surname', 'Last Name', 'LastName', 'Learner Surname', 'Learner Last Name'
    ])

    return (String(first || '') + ' ' + String(last || '')).trim()
  }

  function nameSimilarityScore(a, b) {
    const x = normalizeInvestigationName(a)
    const y = normalizeInvestigationName(b)

    if (!x || !y) return 0
    if (x === y) return 1

    const xp = x.split(' ').filter(Boolean)
    const yp = y.split(' ').filter(Boolean)

    let common = 0
    xp.forEach(part => {
      if (yp.includes(part)) common++
    })

    const tokenScore = common / Math.max(xp.length, yp.length, 1)

    const short = x.length <= y.length ? x : y
    const long = x.length > y.length ? x : y
    const prefixScore = long.startsWith(short.slice(0, Math.min(5, short.length))) ? 0.65 : 0

    return Math.max(tokenScore, prefixScore)
  }

  function findNameInvestigationRows(rows) {
    const headers = getInvestigationHeaders(rows)
    const exactGroups = {}
    const matched = new Set()
    let groupCounter = 1

    rows.forEach(row => {
      const name = normalizeInvestigationName(getInvestigationFullNameValue(row, headers))
      if (!name || name.length < 3) return

      if (!exactGroups[name]) exactGroups[name] = []
      exactGroups[name].push(row)
    })

    Object.keys(exactGroups).forEach(key => {
      const group = exactGroups[key]
      if (group.length < 2) return

      const groupName = 'Name Group ' + groupCounter++

      group.forEach(row => {
        matched.add(row)
        if (!row.duplicateInvestigationGroup) row.duplicateInvestigationGroup = groupName
        if (!row.duplicateInvestigationGroupReason) row.duplicateInvestigationGroupReason = 'Exact name match'
        markInvestigationRow(row, 'name', 'exact', 'Name & Surname exact match')
      })
    })

    const candidates = rows.filter(row => {
      const name = normalizeInvestigationName(getInvestigationFullNameValue(row, headers))
      return name.length >= 5 && !matched.has(row)
    })

    for (let i = 0; i < candidates.length; i++) {
      const group = [candidates[i]]
      const baseName = getInvestigationFullNameValue(candidates[i], headers)

      for (let j = i + 1; j < candidates.length; j++) {
        const compareName = getInvestigationFullNameValue(candidates[j], headers)
        if (nameSimilarityScore(baseName, compareName) >= 0.75) {
          group.push(candidates[j])
        }
      }

      if (group.length < 2) continue

      const groupName = 'Name Group ' + groupCounter++

      group.forEach(row => {
        matched.add(row)
        if (!row.duplicateInvestigationGroup) row.duplicateInvestigationGroup = groupName
        if (!row.duplicateInvestigationGroupReason) row.duplicateInvestigationGroupReason = 'Similar name match'
        markInvestigationRow(row, 'name', 'near', 'Similar name & surname')
      })
    }

    return rows
      .filter(row => matched.has(row))
      .sort((a, b) => {
        const ag = a.duplicateInvestigationGroup || ''
        const bg = b.duplicateInvestigationGroup || ''
        if (ag !== bg) return ag.localeCompare(bg, undefined, { numeric: true })

        const an = normalizeInvestigationName(getInvestigationFullNameValue(a, headers))
        const bn = normalizeInvestigationName(getInvestigationFullNameValue(b, headers))
        return an.localeCompare(bn)
      })
  }




  function addProvinceSupportWithinExistingGroups(rows) {
    const headers = getInvestigationHeaders(rows)
    const grouped = {}

    rows.forEach(row => {
      const group = row.duplicateInvestigationGroup || ''
      if (!group) return
      if (!grouped[group]) grouped[group] = []
      grouped[group].push(row)
    })

    Object.keys(grouped).forEach(groupName => {
      const groupRows = grouped[groupName]
      const provinceGroups = {}

      groupRows.forEach(row => {
        const province = normalizeProvinceValue(getInvestigationProvinceValue(row, headers))
        if (!province || province.length < 3) return
        if (!provinceGroups[province]) provinceGroups[province] = []
        provinceGroups[province].push(row)
      })

      Object.keys(provinceGroups).forEach(province => {
        const provinceRows = provinceGroups[province]
        if (provinceRows.length < 2) return

        provinceRows.forEach(row => {
          markInvestigationRow(row, 'province', 'exact', 'Same province')
        })
      })
    })
  }

  function keepOnlyCompleteInvestigationGroups(rows) {
    const groups = {}

    rows.forEach(row => {
      const group = row.duplicateInvestigationGroup || ''
      if (!group) return
      if (!groups[group]) groups[group] = []
      groups[group].push(row)
    })

    return rows
      .filter(row => {
        const group = row.duplicateInvestigationGroup || ''
        return group && groups[group] && groups[group].length > 1
      })
      .sort((a, b) => {
        const ag = a.duplicateInvestigationGroup || ''
        const bg = b.duplicateInvestigationGroup || ''
        if (ag !== bg) return ag.localeCompare(bg)

        const ar = Number(a['Real Excel Row'] || a['Original Excel Row'] || 0)
        const br = Number(b['Real Excel Row'] || b['Original Excel Row'] || 0)
        return ar - br
      })
  }

  function applyDuplicateInvestigationFilters() {
    const status = document.getElementById('duplicateInvestigationStatus')

    if (!workingRows || !workingRows.length) {
      if (status) status.textContent = 'Upload Excel first before running duplicate checks.'
      return
    }

    if (!duplicateInvestigationBaseRows.length) {
      duplicateInvestigationBaseRows = (allWorkbookRows && allWorkbookRows.length)
        ? allWorkbookRows.slice()
        : workingRows.slice()
    }

    let rows = duplicateInvestigationBaseRows.slice()
    clearDuplicateInvestigationFlags(rows)

    const primaryFiltersActive = duplicateInvestigationFilters.id || duplicateInvestigationFilters.name
    const supportFiltersActive = duplicateInvestigationFilters.cellphone || duplicateInvestigationFilters.province

    if (duplicateInvestigationFilters.id) {
      rows = findIdInvestigationRows(rows)
    }

    if (duplicateInvestigationFilters.name) {
      rows = findNameInvestigationRows(rows)
    }

    if (!primaryFiltersActive && duplicateInvestigationFilters.cellphone) {
      rows = findCellphoneInvestigationRows(rows)
    } else if (duplicateInvestigationFilters.cellphone) {
      findCellphoneInvestigationRows(duplicateInvestigationBaseRows.slice())
    }

    if (!primaryFiltersActive && duplicateInvestigationFilters.province) {
      rows = findProvinceInvestigationRows(rows)
    } else if (duplicateInvestigationFilters.province && rows.length) {
      addProvinceSupportWithinExistingGroups(rows)
    }

    if (primaryFiltersActive && supportFiltersActive && !rows.length) {
      rows = duplicateInvestigationBaseRows.filter(row => row.duplicateInvestigationActive)
    }

    rows = keepOnlyCompleteInvestigationGroups(rows)

    workingRows = rows

    renderTable()
    updateSummary()
    updateAutoFixSuggestions()

    if (status) {
      const activeChecks = []
      if (duplicateInvestigationFilters.id) activeChecks.push('ID Number')
      if (duplicateInvestigationFilters.cellphone) activeChecks.push('Cellphone')
      if (duplicateInvestigationFilters.province) activeChecks.push('Province')
      if (duplicateInvestigationFilters.name) activeChecks.push('Name & Surname')

      if (!activeChecks.length) {
        status.textContent = 'No checks selected. All rows remain visible.'
      } else if (!rows.length) {
        status.textContent = 'No potential matches found for: ' + activeChecks.join(' + ') + '. Try another check or reset the view.'
      } else {
        status.textContent = rows.length + ' possible record(s) found using: ' + activeChecks.join(' + ') + '. Highlighted rows are warnings for manual verification.'
      }
    }
  }

  function resetDuplicateInvestigationFilters() {
    const status = document.getElementById('duplicateInvestigationStatus')

    if (duplicateInvestigationBaseRows.length) {
      workingRows = duplicateInvestigationBaseRows.slice()
      clearDuplicateInvestigationFlags(workingRows)

      renderTable()
      updateSummary()
      updateAutoFixSuggestions()
    }

    if (status) status.textContent = 'Reset complete. All rows are visible again.'
  }


  function renderSimilarityModeControls() {
    const reviewHead = document.querySelector('.review-head')
    if (!reviewHead) return

    let wrap = document.getElementById('similarityModeWrap')

    if (!wrap) {
      reviewHead.insertAdjacentHTML('beforeend', `
        <div id="similarityModeWrap" class="similarity-mode-wrap">
        </div>
      `)
      wrap = document.getElementById('similarityModeWrap')
    }

    const modes = [
      { key: 'auto', label: 'Smart duplicates' },
      { key: 'id', label: 'Show by ID' },
      { key: 'dob', label: 'Show by Date of Birth' },
      { key: 'cell', label: 'Show by Cellphone' }
    ]

    wrap.innerHTML =
      '<span class="cleanup-filter-label">Duplicate view:</span>' +
      modes.map(mode => {
        const active = duplicateViewMode === mode.key ? ' active' : ''
        return `<button type="button" class="similarity-mode-btn${active}" data-mode="${mode.key}">${mode.label}</button>`
      }).join('')

    wrap.querySelectorAll('.similarity-mode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        duplicateViewMode = btn.getAttribute('data-mode') || 'auto'

        wrap.querySelectorAll('.similarity-mode-btn').forEach(b => b.classList.remove('active'))
        btn.classList.add('active')

        const headers = workingRows[0] && workingRows[0].headers ? workingRows[0].headers : []
        if (headers.length) {
          if (allWorkbookRows.length) {
            allWorkbookRows = applyDuplicateFlags(allWorkbookRows, getHeadersForAllTabs(allWorkbookRows), activeRule)
            applyCleanupSheetFilter()
          } else {
            workingRows = applyDuplicateFlags(workingRows, headers, activeRule)
            renderTable()
            updateSummary()
            updateAutoFixSuggestions()
          }
        }
      })
    })
  }

  function normalizeDigits(value) {
    return String(value || '').replace(/\D/g, '')
  }

  function normalizeNameValue(value) {
    return String(value || '').toLowerCase().replace(/[^a-z]/g, '').trim()
  }

  function normalizeDobValue(value) {
    const raw = String(value || '').trim()
    const digits = normalizeDigits(raw)

    if (!raw) return ''

    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw.replace(/\D/g, '')

    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(raw)) {
      const parts = raw.split('/')
      const month = parts[0].padStart(2, '0')
      const day = parts[1].padStart(2, '0')
      let year = parts[2]
      if (year.length === 2) year = Number(year) > 30 ? '19' + year : '20' + year
      return `${year}${month}${day}`
    }

    if (digits.length >= 8) return digits.slice(0, 8)
    if (digits.length === 6) return digits

    return digits
  }

  function getIdentityFieldValue(row, headers, mode) {
    if (mode === 'id') {
      return valueForHeader(row, headers, [
        'ID Number', 'ID No', 'ID', 'Identity Number', 'ID/Passport Number',
        'ID No./Date of Birth', 'ID Number / Date of Birth', 'ID Number / DOB',
        'Learner ID', 'Learner ID Number'
      ])
    }

    if (mode === 'dob') {
      const dob = valueForHeader(row, headers, [
        'Date of Birth', 'DOB', 'Birth Date', 'D.O.B', 'ID No./Date of Birth',
        'ID Number / Date of Birth', 'ID Number / DOB', 'Learner DOB'
      ])

      if (dob) return dob

      const id = valueForHeader(row, headers, [
        'ID Number', 'ID No', 'ID', 'Identity Number', 'ID/Passport Number',
        'ID Number / DOB', 'Learner ID', 'Learner ID Number'
      ])

      const idDigits = normalizeDigits(id)
      return idDigits.length >= 6 ? idDigits.slice(0, 6) : ''
    }

    if (mode === 'cell') {
      return valueForHeader(row, headers, [
        'Cellphone', 'Cellphone Number', 'Cell Phone', 'Cell Number',
        'Cell', 'Phone', 'Phone Number', 'Contact Number', 'Mobile', 'Mobile Number',
        'Learner Cellphone', 'Learner Cellphone Number', 'Learner Contact Number'
      ])
    }

    return ''
  }

  function getRowNameSignal(row, headers) {
    const name = normalizeNameValue(valueForHeader(row, headers, ['Name', 'Full Name', 'Participant Name', 'First Name', 'Learner Name', 'Learner First Name']))
    const surname = normalizeNameValue(valueForHeader(row, headers, ['Surname', 'Last Name', 'Learner Surname', 'Learner Last Name']))
    return `${name}|${surname}`
  }

  function similarityGroupKey(row, headers, mode) {
    const value = getIdentityFieldValue(row, headers, mode)
    const digits = normalizeDigits(value)
    const dob = normalizeDobValue(value)
    const names = getRowNameSignal(row, headers)

    if (mode === 'id') {
      if (digits.length >= 13) return { key: digits, level: 'exact', label: 'Exact ID number match' }
      if (digits.length >= 10) return { key: digits.slice(0, 10), level: 'high', label: 'ID looks similar - last digits differ' }
      if (digits.length >= 6) return { key: digits.slice(0, 6) + '|' + names, level: 'medium', label: 'ID/DOB portion and name look similar' }
    }

    if (mode === 'dob') {
      if (dob.length >= 8) return { key: dob, level: 'exact', label: 'Same date of birth' }
      if (dob.length >= 6) return { key: dob.slice(0, 6) + '|' + names, level: 'medium', label: 'Date of birth and name look similar' }
    }

    if (mode === 'cell') {
      if (digits.length >= 10) return { key: digits, level: 'exact', label: 'Exact cellphone match' }
      if (digits.length >= 6) return { key: digits.slice(0, 6), level: 'high', label: 'Cellphone looks similar - last digits differ' }
      if (digits.length >= 4) return { key: digits.slice(0, 4) + '|' + names, level: 'medium', label: 'Cellphone prefix and name look similar' }
    }

    return null
  }

  function applySimilarityViewFlags(rows, headers) {
    rows.forEach(row => {
      row.similarityMode = ''
      row.similarityLevel = ''
      row.similarityReason = ''
      row.similarityGroupKey = ''
      row.isSimilarityMatch = false
    })

    if (!rows.length || duplicateViewMode === 'auto') return rows

    const groups = {}

    rows.forEach(row => {
      if (row.rowState === 'removed') return

      const group = similarityGroupKey(row, headers, duplicateViewMode)
      if (!group || !group.key) return

      const key = duplicateViewMode + ':' + group.key

      if (!groups[key]) {
        groups[key] = {
          key,
          level: group.level,
          label: group.label,
          rows: []
        }
      }

      groups[key].rows.push(row)
    })

    Object.keys(groups).forEach(key => {
      const group = groups[key]
      if (group.rows.length < 2) return

      group.rows.forEach(row => {
        row.isSimilarityMatch = true
        row.similarityMode = duplicateViewMode
        row.similarityLevel = group.level
        row.similarityReason = group.label
        row.similarityGroupKey = key
        row.isDuplicate = true
        row.groupSize = Math.max(row.groupSize || 0, group.rows.length)

        if (!row.duplicateLevel || row.duplicateLevel === 'Kept original') {
          row.duplicateLevel = group.level === 'exact'
            ? 'Strong duplicate'
            : group.level === 'high'
              ? 'Likely duplicate'
              : 'Possible duplicate'
          row.duplicateReason = group.label
          row.duplicateKey = key
          row._duplicateKey = key
        }
      })
    })

    return rows
  }

  function getHeadersForSelectedTab(rows) {
    if (!rows.length) return []

    const firstSheetHeaders = rows.find(row => Array.isArray(row.sheetHeaders) && row.sheetHeaders.length)

    if (firstSheetHeaders) {
      return firstSheetHeaders.sheetHeaders.filter(h => h !== 'Actions')
    }

    const headers = ['Source Sheet']

    rows.forEach(row => {
      Object.keys(row).forEach(key => {
        if (key.startsWith('_')) return
        if (key === 'rowState') return
        if (key === 'headers') return
        if (key === 'sheetHeaders') return
        if (key === 'Actions') return

        const hasValue = rows.some(r => String(r[key] || '').trim() !== '')
        if (hasValue && !headers.includes(key)) headers.push(key)
      })
    })

    return headers
  }

  function getHeadersForAllTabs(rows) {
    const headers = ['Real Excel Row', 'Source Sheet']

    rows.forEach(row => {
      const sourceHeaders = Array.isArray(row.sheetHeaders) && row.sheetHeaders.length
        ? row.sheetHeaders
        : Object.keys(row)

      sourceHeaders.forEach(key => {
        if (key.startsWith('_')) return
        if (key === 'rowState') return
        if (key === 'headers') return
        if (key === 'sheetHeaders') return
        if (key === 'Actions') return
        if (!headers.includes(key)) headers.push(key)
      })
    })

    return headers
  }


  function applyCleanupSheetFilter() {
    if (!allWorkbookRows.length) return

    workingRows = activeSheetFilter === 'All'
      ? allWorkbookRows.slice()
      : allWorkbookRows.filter(row => row['Source Sheet'] === activeSheetFilter)

    if (activeSheetFilter !== 'All') {
      const tabHeaders = getHeadersForSelectedTab(workingRows)
      workingRows = workingRows.map(row => {
        row.headers = tabHeaders.slice()
        return row
      })
    } else {
      const allHeaders = getHeadersForAllTabs(allWorkbookRows)
      workingRows = workingRows.map(row => {
        row.headers = allHeaders.slice()
        return row
      })
    }

    renderTable()
    updateSummary()
    updateAutoFixSuggestions()
  }

  function loadAllSheets() {
    currentSheetName = 'All Workbook Tabs'

    let allRows = []
    let allHeaders = ['Real Excel Row', 'Source Sheet']

    workbookMeta.sheetNames.forEach(sheetName => {
      const result = rowsFromWorksheet(sheetName)

      result.headers.forEach(h => {
        if (!allHeaders.includes(h)) allHeaders.push(h)
      })

      allRows = allRows.concat(result.rows)
    })

    allRows = allRows.map(row => {
      allHeaders.forEach(h => {
        if (!(h in row)) row[h] = ''
      })
      row.headers = allHeaders.slice()
      return row
    })

    activeRule = resolveRule(allHeaders, 'All Workbook Tabs')
    updateRuleBadge(activeRule)

    allWorkbookRows = applyDuplicateFlags(allRows, allHeaders, activeRule)
    activeSheetFilter = 'All'
    workingRows = allWorkbookRows.slice()
    renderCleanupSheetFilters()
    renderDuplicateInvestigationPanel()
    renderTable()
    updateSummary()
    updateAutoFixSuggestions()

    statusText.innerText = `All workbook tabs loaded: ${workingRows.length} rows from ${workbookMeta.sheetNames.length} sheets.`
  }

  function loadSheet(sheetName) {
    if (!workbookRef || !sheetName) return

    if (sheetName === '__ALL_SHEETS__') {
      loadAllSheets()
      return
    }

    currentSheetName = sheetName

    const result = rowsFromWorksheet(sheetName)
    const headers = result.headers
    const rows = result.rows

    if (!rows.length) {
      workingRows = []
      tableBody.innerHTML = ''
      table.querySelector('thead').innerHTML = ''
      summaryText.innerText = 'No usable rows found'
      updateRuleBadge('generic')
      return
    }

    activeRule = resolveRule(headers, sheetName)
    updateRuleBadge(activeRule)

    allWorkbookRows = applyDuplicateFlags(rows, headers, activeRule)
    activeSheetFilter = 'All'
    workingRows = allWorkbookRows.slice()
    renderCleanupSheetFilters()
    renderSimilarityModeControls()
    renderTable()
    updateSummary()
    updateAutoFixSuggestions()

    statusText.innerText = `${sheetName} loaded: ${workingRows.length} rows.`
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
    const values = headers
      .filter(h => h !== 'Source Sheet')
      .map(h => String(row[h] || '').trim())
      .filter(Boolean)

    return values.length < 2
  }


  function cleanDuplicateValue(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '')
      .trim()
  }

  function cleanPhone(value) {
    return String(value || '').replace(/[^0-9]/g, '').slice(-9)
  }

  function valueForHeader(row, headers, candidates) {
    const key = findHeader(headers, candidates)
    return key ? row[key] : ''
  }

  function buildSmartDuplicateKeys(row, headers) {
    const keys = []

    const name = cleanDuplicateValue(valueForHeader(row, headers, ['Name', 'Full Name', 'Participant Name']))
    const surname = cleanDuplicateValue(valueForHeader(row, headers, ['Surname', 'Last Name']))
    const id = cleanDuplicateValue(valueForHeader(row, headers, ['ID Number', 'ID No', 'ID No./Date of Birth', 'Identity Number', 'Date of Birth']))
    const dob = cleanDuplicateValue(valueForHeader(row, headers, ['Date of Birth', 'DOB']))
    const year = cleanDuplicateValue(valueForHeader(row, headers, ['Year of Birth', 'Birth Year']))
    const phone = cleanPhone(valueForHeader(row, headers, ['Phone', 'Cell', 'Cell Number', 'Contact Number', 'Mobile', 'Mobile Number']))

    if (id && id.length >= 6) keys.push('strong:id:' + id)
    if (dob && dob.length >= 6) keys.push('strong:dob:' + dob)
    if (phone && phone.length >= 8) keys.push('strong:phone:' + phone)

    if (name && surname && dob) keys.push('likely:name-surname-dob:' + name + ':' + surname + ':' + dob)
    if (name && surname && year) keys.push('likely:name-surname-year:' + name + ':' + surname + ':' + year)
    if (name && surname && phone) keys.push('likely:name-surname-phone:' + name + ':' + surname + ':' + phone)

    if (name && surname) keys.push('possible:name-surname:' + name + ':' + surname)

    return keys
  }


  function applyDuplicateFlags(rows, headers, rule) {
    const seen = new Map()

    const prepared = rows.map(row => {
      row.duplicateKey = ''
      row.duplicateLevel = ''
      row.duplicateReason = ''
      row.duplicateFields = []
      row.isDuplicate = false
      row.rowState = row.rowState || 'active'
      return row
    })

    prepared.forEach(row => {
      const smartKeys = buildSmartDuplicateKeys(row, headers)
      const fallbackKey = buildDuplicateKey(row, headers, rule)

      const keys = smartKeys.length ? smartKeys : (fallbackKey ? ['fallback:' + fallbackKey] : [])

      for (const key of keys) {
        if (!key) continue

        if (!seen.has(key)) {
          seen.set(key, row)
          continue
        }

        const first = seen.get(key)
        const level = key.startsWith('strong:') ? 'Strong duplicate'
          : key.startsWith('likely:') ? 'Likely duplicate'
          : key.startsWith('possible:') ? 'Possible duplicate'
          : 'Duplicate'

        row.isDuplicate = true
        row.duplicateKey = key
        row.duplicateLevel = level
        row.duplicateReason = level + ' detected using ' + key.split(':')[1].replace(/-/g, ' ')
        row.duplicateFields = duplicateFieldsFromKey(key, headers)
        row.rowState = 'active'

        if (!first.duplicateLevel) {
          first.duplicateLevel = 'Kept original'
          first.duplicateReason = 'Original record kept for duplicate group'
        }

        break
      }
    })

    return prepared
  }


  function duplicateFieldsFromKey(key, headers) {
    const type = String(key || '').split(':')[1] || ''
    const fields = []

    function add(candidates) {
      const h = findHeader(headers, candidates)
      if (h && !fields.includes(h)) fields.push(h)
    }

    if (type.includes('id')) {
      add(['ID Number', 'ID No', 'ID No./Date of Birth', 'Identity Number'])
    }

    if (type.includes('dob')) {
      add(['Date of Birth', 'DOB', 'ID No./Date of Birth'])
    }

    if (type.includes('phone')) {
      add(['Phone', 'Cell', 'Cell Number', 'Contact Number', 'Mobile', 'Mobile Number'])
    }

    if (type.includes('name')) {
      add(['Name', 'Full Name', 'Participant Name'])
      add(['Surname', 'Last Name'])
    }

    if (type.includes('year')) {
      add(['Year of Birth', 'Birth Year'])
    }

    return fields
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


    if (riskOnlyToggle && riskOnlyToggle.checked) {
      rows = rows.filter(r => {
        const hasEmpty = Object.values(r).some(v => !v || v.toString().trim() === '')
        const idField = Object.keys(r).find(k => k.toLowerCase().includes('id'))
        let shortId = false
        if (idField) {
          const val = (r[idField] || '').toString().replace(/\D/g,'')
          shortId = val && val.length < 10
        }
        return hasEmpty || shortId
      })
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

      if (row.isSimilarityMatch) {
        tr.classList.add('row-similarity')
        tr.classList.add('row-similarity-' + (row.similarityLevel || 'medium'))
      }

      if (row.duplicateInvestigationActive && row.duplicateInvestigationTags && row.duplicateInvestigationTags.includes('id')) {
        tr.classList.add('row-investigation-id')
        if (row.duplicateInvestigationLevels && row.duplicateInvestigationLevels.includes('near')) {
          tr.classList.add('row-investigation-id-near')
        }
      }

      if (row.duplicateInvestigationActive && row.duplicateInvestigationTags && row.duplicateInvestigationTags.includes('cellphone')) {
        tr.classList.add('row-investigation-cellphone')
        if (row.duplicateInvestigationLevels && row.duplicateInvestigationLevels.includes('near')) {
          tr.classList.add('row-investigation-cellphone-near')
        }
      }

      if (row.duplicateInvestigationActive && row.duplicateInvestigationTags && row.duplicateInvestigationTags.includes('province')) {
        tr.classList.add('row-investigation-province')
      }

      if (row.duplicateInvestigationActive && row.duplicateInvestigationTags && row.duplicateInvestigationTags.includes('name')) {
        tr.classList.add('row-investigation-name')
      }


      // ===== RISK FLAGS =====
      const hasEmptyCell = headers.some(h => !row[h] || row[h].toString().trim() === '')
      const shortId = (function () {
        const idKey = headers.find(h => h.toLowerCase().includes('id'))
        if (!idKey) return false
        const val = (row[idKey] || '').toString().replace(/\D/g,'')
        return val && val.length < 10
      })()

      if (hasEmptyCell) {
        tr.classList.add('row-warning')
      }

      if (shortId) {
        tr.classList.add('row-risk')
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
      return `<span class="status-pill status-removed">Marked for review</span>`
    }

    if (row.keepChoice) {
      return `<span class="status-pill status-kept">Kept</span>`
    }

    if (row.duplicateInvestigationActive && row.duplicateInvestigationReasons && row.duplicateInvestigationReasons.length) {
      const sourceRow = row.__excelRowNumber || row.__rowIndex || row.rowNumber || row._rowNumber || row.excelRowNumber || ''
      const sourceSheet = row.__sourceSheet || row['Source Sheet'] || ''
      const rowText = sourceRow ? `<div class="status-row-ref">Actual Excel row: ${sourceRow}${sourceSheet ? ' • Sheet: ' + sourceSheet : ''}</div>` : ''
      return `<span class="status-pill status-duplicate">${row.duplicateInvestigationReasons.join(' + ')}</span>${rowText}`
    }

    if (row.isSimilarityMatch) {
      const label = row.similarityLevel === 'exact' ? 'Exact match'
        : row.similarityLevel === 'high' ? 'High similarity'
        : 'Possible similarity'
      return `<span class="status-pill status-duplicate">${label} (${row.groupSize})</span>`
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
          <button class="row-btn remove" data-action="remove" data-id="${row._id}">Mark for Review</button>
        </div>
      `
    }

    return `
      <div class="row-actions">
        <button class="row-btn remove" data-action="remove" data-id="${row._id}">Mark for Review</button>
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
          row.rowState = 'active'
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

    summaryText.innerText = `${total} rows loaded • ${duplicates} duplicate rows flagged • ${kept} marked for review • ${visible} currently visible`
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


function updateAutoFixSuggestions() {
  const container = document.getElementById('autoFixContent')
  if (!container) return

  if (!workingRows || !workingRows.length) {
    container.innerHTML = 'No suggestions yet'
    return
  }

  let missingCount = 0
  let shortIdCount = 0

  workingRows.forEach(r => {
    const values = Object.values(r)

    if (values.some(v => !v || v.toString().trim() === '')) {
      missingCount++
    }

    const idField = Object.keys(r).find(k => k.toLowerCase().includes('id'))
    if (idField) {
      const val = (r[idField] || '').toString().replace(/\D/g,'')
      if (val && val.length < 10) {
        shortIdCount++
      }
    }
  })

  const suggestions = []

  if (missingCount > 0) {
    suggestions.push(`⚠️ ${missingCount} rows have missing fields`)
  }

  if (shortIdCount > 0) {
    suggestions.push(`🪪 ${shortIdCount} rows may have invalid ID numbers`)
  }

  if (!suggestions.length) {
    container.innerHTML = '✅ Data looks clean'
    return
  }

  container.innerHTML = suggestions.map(s => `<div class="auto-fix-item">${s}</div>`).join('')
}


document.addEventListener('DOMContentLoaded', function () {
  const btn = document.getElementById('autoCleanBtn')
  if (!btn) return

  btn.addEventListener('click', function () {
    if (!workingRows || !workingRows.length) return

    // trim values
    workingRows.forEach(r => {
      Object.keys(r).forEach(k => {
        if (typeof r[k] === 'string') {
          r[k] = r[k].trim()
        }
      })
    })

    // remove fully empty rows
    workingRows = workingRows.filter(r => {
      const values = Object.values(r)
      return values.some(v => v && v.toString().trim() !== '')
    })

    renderTable()
    updateSummary()
    if (typeof updateAutoFixSuggestions === 'function') {
      updateAutoFixSuggestions()
    }
  })
})


document.addEventListener('DOMContentLoaded', () => {
  const style = document.createElement('style')
  style.innerHTML = `
    .cleanup-sheet-filter-wrap{
      display:flex;
      gap:10px;
      align-items:center;
      flex-wrap:wrap;
      justify-content:flex-end;
      margin-left:auto;
      max-width:760px;
    }
    .cleanup-filter-label{
      font-weight:900;
      color:#374151;
    }
    .cleanup-sheet-filter{
      border:1px solid #e5e7eb;
      background:#fff;
      color:#374151;
      border-radius:999px;
      padding:9px 13px;
      font-weight:900;
      cursor:pointer;
    }
    .cleanup-sheet-filter.active{
      background:#111827;
      color:#fff;
      border-color:#111827;
      box-shadow:0 0 0 3px rgba(14,165,233,.18);
    }
    @media(max-width:900px){
      .cleanup-sheet-filter-wrap{
        justify-content:flex-start;
        width:100%;
        margin-top:12px;
      }
    }
  `
  document.head.appendChild(style)
})

window.pmCleanupSheetFilterInstalled = true


window.pmCleanupTabSpecificColumnsPatch = true


window.pmCleanupDateFormatPatch = true


window.pmCleanupDuplicateLogicPatch = true


document.addEventListener('DOMContentLoaded', () => {
  const style = document.createElement('style')
  style.innerHTML = `
    .duplicate-cell-alert{
      background:#fff1f2 !important;
      box-shadow:inset 0 0 0 2px #fb7185;
    }
    .duplicate-cell-alert input{
      background:#fff1f2 !important;
      border-color:#fb7185 !important;
      color:#7f1d1d !important;
      font-weight:900;
    }
  `
  document.head.appendChild(style)
})

window.pmCleanupDuplicateCellHighlightPatch = true
