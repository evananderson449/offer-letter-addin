/**
 * document-ops.js
 * Utility functions for Handl Health Offer Letter Generator Add-in
 * Status display and initialization only — all document manipulation
 * is handled in-memory by taskpane.js using JSZip.
 */
console.log('[document-ops.js] v4.0 loaded');

/**
 * Display status message to user
 * @param {string} message - Status message to display
 * @param {string} type - 'success' | 'error' | 'info'
 */
function showStatus(message, type = 'info') {
  console.log(`[${type.toUpperCase()}] ${message}`);

  const statusElement = document.getElementById('status-message');
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = 'status-message status-' + type;
    statusElement.style.display = 'block';

    // Auto-hide after 4 seconds for success/info
    if (type === 'success' || type === 'info') {
      setTimeout(() => {
        statusElement.style.display = 'none';
      }, 4000);
    }
  }
}

/**
 * Initialize Office.js add-in
 */
window.initializeAddIn = async function () {
  try {
    if (typeof Word === 'undefined') {
      throw new Error('Office.js (Word) is not loaded.');
    }
    console.log('Add-in initialized');
    return true;
  } catch (error) {
    console.error('Initialization error:', error);
    showStatus('Failed to initialize add-in', 'error');
    return false;
  }
};
