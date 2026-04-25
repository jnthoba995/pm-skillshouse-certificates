document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('registerFile')
  const fileName = document.getElementById('registerFileName')
  const status = document.getElementById('registerStatus')
  const processBtn = document.getElementById('processRegisterBtn')
  const summary = document.getElementById('registerSummary')

  if (fileInput && fileName) {
    fileInput.addEventListener('change', function () {
      const file = fileInput.files && fileInput.files[0]
      fileName.innerText = file ? file.name : 'No file chosen'
      status.innerText = file ? 'Register selected. Ready to process.' : 'No register processed yet'
    })
  }

  if (processBtn) {
    processBtn.addEventListener('click', function () {
      status.innerText = 'Register processing workflow ready. OCR connection comes next.'
      summary.innerText = 'Awaiting captured rows'
    })
  }
})
