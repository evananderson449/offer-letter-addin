/**
 * document-ops.js
 * Core Office.js logic for Handl Health Offer Letter Generator Add-in
 * Handles placeholder replacement, bonus paragraph deletion, and file save/rename
 */

/**
 * Display status message to user (UI integration point)
 * @param {string} message - Status message to display
 * @param {string} type - 'success' | 'error' | 'info'
 */
function showStatus(message, type = 'info') {
  console.log(`[${type.toUpperCase()}] ${message}`);

  // UI integration: Update taskpane status element if available
  const statusElement = document.getElementById('status-message');
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `status-${type}`;
    statusElement.style.display = 'block';

    // Auto-hide success messages after 3 seconds
    if (type === 'success') {
      setTimeout(() => {
        statusElement.style.display = 'none';
      }, 3000);
    }
  }
}

/**
 * Main document update function - searches and replaces all placeholders
 * @param {Object} formData - Form values from taskpane
 * @param {string} formData.name - Full name
 * @param {string} formData.title - Job title
 * @param {string} formData.startDate - Formatted start date (Month DD, YYYY)
 * @param {string} formData.supervisor - Supervisor name
 * @param {string} formData.salary - Formatted salary with $ sign
 * @param {boolean} formData.bonusEnabled - Whether bonus section is included
 * @param {string} formData.bonusPctRange - e.g. "10-20"
 * @param {string} formData.bonusDollarRange - e.g. "$12,000-$24,000"
 * @param {string} formData.exempt - "Exempt" or "Non-Exempt"
 * @param {string} formData.sharesNum - Number of shares (formatted with commas)
 * @param {string} formData.sharesPct - Shares percentage
 * @param {string} formData.sharesValue - Dollar value of shares (formatted)
 * @param {string} formData.expirationDate - Formatted expiration date (Month DD, YYYY)
 */
window._updateDocumentCore = async function(formData) {
  try {
    showStatus('Generating offer letter...', 'info');

    await Word.run(async (context) => {
      const body = context.document.body;

      // 1. Replace all placeholder text
      // Map of placeholder strings to replacement values
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

      // Process each replacement
      for (const { find, replace } of replacements) {
        const searchResults = body.search(find, { matchCase: true, matchWholeWord: false });
        searchResults.load('items');
        await context.sync();

        // Replace all occurrences of this placeholder
        for (let i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].insertText(replace, Word.InsertLocation.replace);
        }
      }

      await context.sync();

      // 2. Handle bonus paragraph deletion if bonus is disabled
      if (!formData.bonusEnabled) {
        await deleteBonusParagraph(context);
      }

      // 3. Handle exempt/non-exempt logic
      await handleExemptLogic(context, formData.exempt);

      await context.sync();
    });

    showStatus('Offer letter generated successfully!', 'success');
  } catch (error) {
    console.error('Error updating document:', error);
    showStatus(`Error: ${error.message}`, 'error');
    throw error;
  }
};

/**
 * Delete the bonus paragraph from the document
 * Finds "Discretionary, performance-based bonus" and removes its paragraph
 * @param {Word.RequestContext} context - Word context
 */
async function deleteBonusParagraph(context) {
  try {
    const body = context.document.body;
    const bonusSearchResults = body.search('Discretionary, performance-based bonus', {
      matchCase: false,
      matchWholeWord: false
    });

    bonusSearchResults.load('items');
    await context.sync();

    // Get parent paragraph of each search result and delete it
    for (let i = 0; i < bonusSearchResults.items.length; i++) {
      const item = bonusSearchResults.items[i];
      const paragraphs = item.paragraphs;
      paragraphs.load('items');
      await context.sync();
      if (paragraphs.items.length > 0) {
        paragraphs.items[0].delete();
      }
    }

    await context.sync();
  } catch (error) {
    console.error('Error deleting bonus paragraph:', error);
    // Non-fatal error - continue with other operations
  }
}

/**
 * Handle FLSA exempt/non-exempt logic
 * If Non-Exempt: Replace "will not be eligible" with "will be eligible"
 * If Exempt: Leave "will not be eligible" as-is
 * @param {Word.RequestContext} context - Word context
 * @param {string} exempt - "Exempt" or "Non-Exempt"
 */
async function handleExemptLogic(context, exempt) {
  try {
    if (exempt === 'Non-Exempt') {
      const body = context.document.body;
      const searchResults = body.search('will not be eligible', {
        matchCase: false,
        matchWholeWord: false
      });

      searchResults.load('items');
      await context.sync();

      // Replace "will not be eligible" with "will be eligible"
      for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].insertText('will be eligible', Word.InsertLocation.replace);
      }

      await context.sync();
    }
  } catch (error) {
    console.error('Error handling exempt logic:', error);
    // Non-fatal error - continue
  }
}

/**
 * Save/rename the document
 * Attempts Office.js save first, then falls back to download
 * @param {string} employeeName - Full name of employee (e.g., "Jane Smith")
 */
window._saveDocumentCore = async function(employeeName) {
  try {
    const filename = `Handl_Offer_Letter_${employeeName.replace(/\s+/g, '_')}.docx`;
    showStatus(`Saving as ${filename}...`, 'info');

    // Attempt 1: Use Office.context.document.saveAsync() for built-in save dialog
    // This provides a "Save As" experience in Word Online
    try {
      await new Promise((resolve, reject) => {
        Office.context.document.saveAsync(function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error.message));
          }
        });
      });

      showStatus(`Document saved as ${filename}`, 'success');
      return;
    } catch (saveError) {
      console.warn('Built-in save failed, attempting download fallback:', saveError);
    }

    // Attempt 2: Fall back to getFileAsync() + browser download
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 262144 }, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          downloadFile(result.value, filename);
          showStatus(`Document downloaded as ${filename}`, 'success');
          resolve();
        } else {
          reject(new Error(`Failed to get file: ${result.error.message}`));
        }
      });
    });
  } catch (error) {
    console.error('Error saving document:', error);
    showStatus(`Error saving document: ${error.message}`, 'error');
    throw error;
  }
};

/**
 * Helper: Download file to user's computer
 * Reconstructs file from slices and triggers browser download
 * @param {Office.File} file - File object from getFileAsync
 * @param {string} filename - Desired filename for download
 */
function downloadFile(file, filename) {
  const sliceCount = file.sliceCount;
  const sliceData = [];

  function readSlice(index) {
    file.getSliceAsync(index, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        sliceData.push(result.value.data);
        if (index + 1 < sliceCount) {
          readSlice(index + 1);
        } else {
          file.closeAsync();
          const blob = new Blob(sliceData, {
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
        }
      }
    });
  }

  readSlice(0);
}

/**
 * Validate form data before processing
 * @param {Object} formData - Form values to validate
 * @returns {Object} { valid: boolean, errors: string[] }
 */
window.validateFormData = function(formData) {
  const errors = [];

  if (!formData.name || formData.name.trim() === '') {
    errors.push('Employee name is required');
  }
  if (!formData.title || formData.title.trim() === '') {
    errors.push('Job title is required');
  }
  if (!formData.startDate || formData.startDate.trim() === '') {
    errors.push('Start date is required');
  }
  if (!formData.supervisor || formData.supervisor.trim() === '') {
    errors.push('Supervisor name is required');
  }
  if (!formData.salary || formData.salary.trim() === '') {
    errors.push('Salary is required');
  }
  if (formData.bonusEnabled) {
    if (!formData.bonusPctRange || formData.bonusPctRange.trim() === '') {
      errors.push('Bonus percentage range is required');
    }
    if (!formData.bonusDollarRange || formData.bonusDollarRange.trim() === '') {
      errors.push('Bonus dollar range is required');
    }
  }
  if (!formData.exempt || (formData.exempt !== 'Exempt' && formData.exempt !== 'Non-Exempt')) {
    errors.push('FLSA status must be "Exempt" or "Non-Exempt"');
  }
  if (!formData.sharesNum || formData.sharesNum.trim() === '') {
    errors.push('Share count is required');
  }
  if (!formData.sharesPct || formData.sharesPct.trim() === '') {
    errors.push('Share percentage is required');
  }
  if (!formData.sharesValue || formData.sharesValue.trim() === '') {
    errors.push('Share value is required');
  }
  if (!formData.expirationDate || formData.expirationDate.trim() === '') {
    errors.push('Expiration date is required');
  }

  return {
    valid: errors.length === 0,
    errors
  };
};

/**
 * Initialize Office.js add-in
 * Call this once when taskpane loads
 */
window.initializeAddIn = async function() {
  try {
    // Require the Office.js library (assumes it's loaded globally)
    if (typeof Word === 'undefined') {
      throw new Error('Office.js (Word) is not loaded. Ensure the script is included in taskpane.html.');
    }

    showStatus('Add-in ready', 'success');
    return true;
  } catch (error) {
    console.error('Initialization error:', error);
    showStatus('Failed to initialize add-in', 'error');
    return false;
  }
};

// Export for module systems if needed
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    updateDocument: window._updateDocumentCore,
    saveDocument: window._saveDocumentCore,
    validateFormData: window.validateFormData,
    initializeAddIn: window.initializeAddIn,
    showStatus
  };
}
