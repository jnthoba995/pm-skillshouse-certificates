document.addEventListener('DOMContentLoaded', function () {
  const expandBtn = document.getElementById('expandReviewBtn')
  const overlay = document.getElementById('reviewOverlay')
  const closeBtn = document.getElementById('closeReviewBtn')

  if (!expandBtn || !overlay || !closeBtn) return

  expandBtn.addEventListener('click', function () {
    overlay.style.display = 'flex'
  })

  closeBtn.addEventListener('click', function () {
    overlay.style.display = 'none'
  })
})
