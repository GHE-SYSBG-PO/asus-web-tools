/* ========================================
   App Navigation & Tool Switching
   ======================================== */
(function () {
  'use strict';

  const sidebar = document.getElementById('sidebar');
  const sidebarToggle = document.getElementById('sidebar-toggle');
  const navItems = document.querySelectorAll('.nav-item[data-tool]');

  // ── Sidebar toggle ────────────────────
  sidebarToggle.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
  });

  // ── Tool switching ────────────────────
  function activateTool(toolId) {
    // Update nav items
    navItems.forEach((item) => {
      const isActive = item.dataset.tool === toolId;
      item.classList.toggle('nav-item--active', isActive);
      item.setAttribute('aria-selected', isActive);
    });

    // Update panels
    document.querySelectorAll('.tool-panel').forEach((panel) => {
      panel.classList.toggle('hidden', panel.id !== `panel-${toolId}`);
    });

    // Save last used tool
    try {
      localStorage.setItem('devtools-active', toolId);
    } catch (_) {}
  }

  navItems.forEach((item) => {
    item.addEventListener('click', () => {
      activateTool(item.dataset.tool);
    });
  });

  // ── Restore last tool ─────────────────
  function init() {
    let lastTool = 'ada-contrast';
    try {
      lastTool = localStorage.getItem('devtools-active') || 'ada-contrast';
    } catch (_) {}

    // Ensure the saved tool exists
    const exists = [...navItems].some((i) => i.dataset.tool === lastTool);
    activateTool(exists ? lastTool : 'ada-contrast');
  }

  init();
})();
