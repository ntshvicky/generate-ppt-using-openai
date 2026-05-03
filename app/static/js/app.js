// Global app utilities

// Auto-dismiss flash messages after 5s
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('[class^="flash-"]').forEach(el => {
    setTimeout(() => el.remove(), 5000);
  });

  // Confirm delete forms
  document.querySelectorAll('form[data-confirm]').forEach(form => {
    form.addEventListener('submit', e => {
      if (!confirm(form.dataset.confirm)) e.preventDefault();
    });
  });

  // Animate stat cards on load
  document.querySelectorAll('.stat-card').forEach((el, i) => {
    el.style.animationDelay = `${i * 60}ms`;
    el.classList.add('animate-fade-in');
  });
});
