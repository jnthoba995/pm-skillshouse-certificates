document.addEventListener('DOMContentLoaded', function () {
  const expandBtn = document.getElementById('expandReviewBtn')
  const overlay = document.getElementById('reviewOverlay')
  const closeBtn = document.getElementById('closeReviewBtn')
  const overlayBody = document.getElementById('reviewOverlayBody')
  const originalParent = document.querySelector('.review-scroll-shell')
  const tableWrap = document.querySelector('.review-scroll-shell .table-wrap')

  if (!expandBtn || !overlay || !closeBtn || !overlayBody || !originalParent || !tableWrap) return

  expandBtn.addEventListener('click', function () {
    overlayBody.appendChild(tableWrap)
    overlay.style.display = 'flex'
  })

  closeBtn.addEventListener('click', function () {
    originalParent.appendChild(tableWrap)
    overlay.style.display = 'none'
  })
})
