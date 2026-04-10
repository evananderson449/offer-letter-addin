/**
 * Handl Offer Letter Generator - Taskpane Integration Layer
 *
 * This file serves as the integration glue between the form UI (form-state.js)
 * and the Word document operations (document-ops.js).
 *
 * Architecture:
 * 1. Office.onReady() is the only entry point
 * 2. Initializes form state management
 * 3. Sets up bridge functions that transform form data into document operations
 * 4. Exports updateDocument() and saveDocument() for form button handlers
 */

/**
 * Office.onReady is the single entry point for the add-in
 * Initializes everything when Office.js is ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log('Handl Offer Letter Generator ready');

    // Initialize form state management
    window.initFormState();

    // Initialize add-in (Office.js validation)
    window.initializeAddIn();
  }
});

/**
 * Bridge function: Transform form data and call document-ops.js
 * This is the "Generate Letter" button handler
 * Validates form, transforms raw form data to document format, and inserts into Word
 */
window.updateDocument = async function() {
  try {
    // Get raw form data from form-state.js
    const raw = window.getFormData();

    // Validate required fields
    if (!window.checkFormStatus()) {
      const statusElement = document.getElementById('status-message');
      if (statusElement) {
        statusElement.textContent = 'Please fill in all required fields';
        statusElement.className = 'status-message status-error';
        statusElement.style.display = 'block';
        setTimeout(() => {
          statusElement.style.display = 'none';
        }, 3000);
      } else {
        alert('Please fill in all required fields');
      }
      return;
    }

    // Transform raw form values into the format expected by document-ops
    const formData = {
      name: raw.f_name,
      title: raw.f_title,
      startDate: window.formatDate(raw.f_start_date),
      supervisor: raw.f_supervisor,
      salary: window.formatCurrency(parseFloat(raw.f_salary) || 0),
      bonusEnabled: raw.bonusToggle,
      bonusPctRange: raw.bonusToggle ? `${raw.f_bonus_pct_a}-${raw.f_bonus_pct_b}` : '',
      bonusDollarRange: raw.bonusToggle ? `${raw.f_bonus_dollar_a}-${raw.f_bonus_dollar_b}` : '',
      exempt: raw.f_exempt === 'exempt' ? 'Exempt' : 'Non-Exempt',
      sharesNum: raw.f_shares_num || '0',
      sharesPct: raw.f_shares_pct || '0',
      sharesValue: raw.f_shares_val || '$0',
      expirationDate: window.formatDate(raw.f_expiration)
    };

    // Call the core document operation (from document-ops.js)
    await window._updateDocumentCore(formData);
  } catch (error) {
    console.error('Error in updateDocument bridge:', error);
    const statusElement = document.getElementById('status-message');
    if (statusElement) {
      statusElement.textContent = `Error: ${error.message}`;
      statusElement.className = 'status-message status-error';
      statusElement.style.display = 'block';
    }
  }
};

/**
 * Bridge function: Extract name from form and trigger save
 * This is the "Save As" button handler
 * Gets employee name from form and calls the save operation
 */
window.saveDocument = async function() {
  try {
    // Get raw form data from form-state.js
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';

    // Call the core save operation (from document-ops.js)
    await window._saveDocumentCore(name);
  } catch (error) {
    console.error('Error in saveDocument bridge:', error);
    const statusElement = document.getElementById('status-message');
    if (statusElement) {
      statusElement.textContent = `Error: ${error.message}`;
      statusElement.className = 'status-message status-error';
      statusElement.style.display = 'block';
    }
  }
};
