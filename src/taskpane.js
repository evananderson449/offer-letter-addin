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
// Debounce timer for live updates
let _liveUpdateTimer = null;
const LIVE_UPDATE_DELAY = 800; // ms after last keystroke

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log('Handl Offer Letter Generator ready');

    // Initialize form state management
    window.initFormState();

    // Initialize add-in (Office.js validation)
    window.initializeAddIn();

    // Attach live update listeners to all form inputs (debounced)
    setupLiveUpdate();
  }
});

/**
 * Attach debounced live update to all form fields
 * Triggers updateDocument() 800ms after the user stops typing
 */
function setupLiveUpdate() {
  const form = document.getElementById('offerForm');
  if (!form) return;

  // Listen on the entire form for input/change events (event delegation)
  form.addEventListener('input', debouncedLiveUpdate);
  form.addEventListener('change', debouncedLiveUpdate);
}

function debouncedLiveUpdate() {
  // Only auto-update if all required fields are filled
  if (!window.checkFormStatus()) return;

  clearTimeout(_liveUpdateTimer);
  _liveUpdateTimer = setTimeout(() => {
    console.log('Live update triggered');
    window.updateDocument();
  }, LIVE_UPDATE_DELAY);
}

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
 * Bridge function: Save the document with proper naming
 * In Word Online, we set the document title/properties and trigger a save,
 * then instruct the user to use File > Save As or download from SharePoint.
 */
window.saveDocument = async function() {
  try {
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';
    const filename = `Handl_Offer_Letter_${name.replace(/\s+/g, '_')}`;

    showStatus(`Saving document...`, 'info');

    // Set the document title property to the desired filename
    await Word.run(async (context) => {
      const properties = context.document.properties;
      properties.title = filename;
      await context.sync();
    });

    // Trigger Office save
    await new Promise((resolve, reject) => {
      Office.context.document.saveAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error ? result.error.message : 'Save failed'));
        }
      });
    });

    showStatus(`Document saved! To download as "${filename}.docx", use File > Save As > Download a Copy.`, 'success');
  } catch (error) {
    console.error('Error saving document:', error);

    // Fallback: try getFileAsync download
    try {
      const raw = window.getFormData();
      const name = raw.f_name || 'Unknown';
      await window._saveDocumentCore(name);
    } catch (fallbackError) {
      showStatus(`To save: use File > Save As > Download a Copy, then rename to "Handl_Offer_Letter_${(raw?.f_name || 'Name').replace(/\s+/g, '_')}.docx"`, 'info');
    }
  }
};
