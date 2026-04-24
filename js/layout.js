async function loadLayout() {
  const headerMount = document.getElementById('appHeader')
  const sidebarMount = document.getElementById('appSidebar')

  if (headerMount) {
    const header = await fetch('components/header.html').then(r => r.text())
    headerMount.innerHTML = header
  }

  if (sidebarMount) {
    const sidebar = await fetch('components/sidebar.html').then(r => r.text())
    sidebarMount.innerHTML = sidebar
  }

  const user = JSON.parse(localStorage.getItem('pm_user') || '{}')
  const currentUser = document.getElementById('currentUser')
  if (user && user.email && currentUser) {
    currentUser.innerText = user.email
  }

  const logoutBtn = document.getElementById('logoutBtn')
  if (logoutBtn) {
    logoutBtn.onclick = () => {
      localStorage.removeItem('pm_user')
      window.location.href = 'login.html'
    }
  }

  const page = window.location.pathname.split('/').pop()
  document.querySelectorAll('.nav-link').forEach(link => {
    const href = link.getAttribute('href')
    if (href === page) link.classList.add('active')
  })
}

document.addEventListener('DOMContentLoaded', loadLayout)
