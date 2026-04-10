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

    // NOTE: Live update is disabled because Generate Letter creates a NEW document.
    // Auto-updating on keystrokes would create many unwanted copies.
    // Users click "Generate Letter" when ready.
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
 * Build the formData object from raw form values
 */
function buildFormData() {
  const raw = window.getFormData();
  return {
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
}

/**
 * Bridge function: Generate Letter
 *
 * CRITICAL WORKFLOW: This creates a NEW document from the template content,
 * applies replacements to the copy, and opens it — leaving the original template intact.
 *
 * Step 1: Get the current document (template) as a Base64-encoded .docx
 * Step 2: Create a new document from that Base64 content
 * Step 3: Apply all placeholder replacements in the new document
 */
window.updateDocument = async function() {
  try {
    // Validate required fields
    if (!window.checkFormStatus()) {
      const statusElement = document.getElementById('status-message');
      if (statusElement) {
        statusElement.textContent = 'Please fill in all required fields';
        statusElement.className = 'status-message status-error';
        statusElement.style.display = 'block';
        setTimeout(() => { statusElement.style.display = 'none'; }, 3000);
      }
      return;
    }

    const formData = buildFormData();
    showStatus('Creating offer letter copy...', 'info');

    // Get the template content as Base64
    const templateBase64 = await new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 262144 }, async function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('Failed to read template: ' + (result.error ? result.error.message : 'Unknown error')));
          return;
        }

        const file = result.value;
        const sliceCount = file.sliceCount;
        const sliceData = [];

        function readNextSlice(index) {
          file.getSliceAsync(index, function(sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              sliceData.push(sliceResult.value.data);
              if (index + 1 < sliceCount) {
                readNextSlice(index + 1);
              } else {
                file.closeAsync();
                // Combine all slices into a single base64 string
                const bytes = new Uint8Array(sliceData.reduce((acc, slice) => {
                  const arr = new Uint8Array(slice);
                  const combined = new Uint8Array(acc.length + arr.length);
                  combined.set(acc);
                  combined.set(arr, acc.length);
                  return combined;
                }, new Uint8Array(0)));
                // Convert to base64
                let binary = '';
                for (let i = 0; i < bytes.length; i++) {
                  binary += String.fromCharCode(bytes[i]);
                }
                resolve(btoa(binary));
              }
            } else {
              file.closeAsync();
              reject(new Error('Failed to read file slice'));
            }
          });
        }
        readNextSlice(0);
      });
    });

    // Create a new document from the template base64
    showStatus('Opening new document with replacements...', 'info');
    await Word.run(async (context) => {
      const newDoc = context.application.createDocument(templateBase64);
      await context.sync();

      // The new document is now the active document
      // Apply all replacements to it
      const body = newDoc.body;

      const replacements = [
        { find: '[NAME]', replace: formData.name },
        { find: '[TITLE]', replace: formData.title },
        { find: '[START DATE]', replace: formData.startDate },
        { find: '[SUPERVISOR]', replace: formData.supervisor },
        { find: '[SALARY]', replace: formData.salary },
        { find: '[BONUS A % - BONUS B %]', replace: formData.bonusPctRange },
        { find: '[BONUS A $ - BONUS B $]', replace: formData.bonusDollarRange },
        { find: '[EXEMPT]', replace: formData.exempt },
        { find: '[# OF SHARES]', replace: formData.sharesNum },
        { find: '[SHARES %]', replace: formData.sharesPct },
        { find: '[$ SHARES]', replace: formData.sharesValue },
        { find: '[EXPIRATION DATE]', replace: formData.expirationDate }
      ];

      for (const { find, replace } of replacements) {
        const searchResults = body.search(find, { matchCase: true, matchWholeWord: false });
        searchResults.load('items');
        await context.sync();
        for (let i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].insertText(replace, Word.InsertLocation.replace);
        }
      }

      // Handle bonus paragraph deletion
      if (!formData.bonusEnabled) {
        const bonusResults = body.search('Discretionary, performance-based bonus', { matchCase: false, matchWholeWord: false });
        bonusResults.load('items');
        await context.sync();
        for (let i = 0; i < bonusResults.items.length; i++) {
          const paragraphs = bonusResults.items[i].paragraphs;
          paragraphs.load('items');
          await context.sync();
          if (paragraphs.items.length > 0) {
            paragraphs.items[0].delete();
          }
        }
      }

      // Handle FLSA exempt logic
      if (formData.exempt === 'Non-Exempt') {
        const exemptResults = body.search('will not be eligible', { matchCase: false, matchWholeWord: false });
        exemptResults.load('items');
        await context.sync();
        for (let i = 0; i < exemptResults.items.length; i++) {
          exemptResults.items[i].insertText('will be eligible', Word.InsertLocation.replace);
        }
      }

      await context.sync();

      // Open the new document (this switches the user to the new doc)
      newDoc.open();
      await context.sync();
    });

    showStatus(`Offer letter created for ${formData.name}! The original template is unchanged.`, 'success');
  } catch (error) {
    console.error('Error generating offer letter:', error);
    showStatus(`Error: ${error.message}`, 'error');
  }
};

/**
 * Bridge function: Save / Download the current document
 * Works after Generate Letter has created a new document.
 * Attempts to download the file with the correct name.
 */
window.saveDocument = async function() {
  try {
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';
    const filename = `Handl_Offer_Letter_${name.replace(/\s+/g, '_')}.docx`;

    showStatus(`Preparing download as ${filename}...`, 'info');

    // Try to download the current document with the correct filename
    await new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 262144 }, function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error ? result.error.message : 'Failed to get file'));
          return;
        }

        const file = result.value;
        const sliceCount = file.sliceCount;
        const sliceData = [];

        function readSlice(index) {
          file.getSliceAsync(index, function(sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              sliceData.push(sliceResult.value.data);
              if (index + 1 < sliceCount) {
                readSlice(index + 1);
              } else {
                file.closeAsync();
                // Combine slices into blob and trigger download
                const bytes = new Uint8Array(sliceData.reduce((acc, slice) => {
                  const arr = new Uint8Array(slice);
                  const combined = new Uint8Array(acc.length + arr.length);
                  combined.set(acc);
                  combined.set(arr, acc.length);
                  return combined;
                }, new Uint8Array(0)));
                const blob = new Blob([bytes], {
                  type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);
                showStatus(`Downloaded as ${filename}`, 'success');
                resolve();
              }
            } else {
              file.closeAsync();
              reject(new Error('Failed to read file slice'));
            }
          });
        }
        readSlice(0);
      });
    });
  } catch (error) {
    console.error('Error downloading document:', error);
    // If programmatic download fails (common in Word Online iframe sandbox),
    // guide the user to download manually
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';
    const filename = `Handl_Offer_Letter_${name.replace(/\s+/g, '_')}.docx`;
    showStatus(`To download: File > Save As > Download a Copy. Rename to "${filename}"`, 'info');
  }
};
