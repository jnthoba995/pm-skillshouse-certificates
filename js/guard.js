(function () {
  const user = localStorage.getItem('pm_user')
  if (!user) {
    window.location.href = 'login.html'
  }
})()
