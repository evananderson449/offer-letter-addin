/**
 * Handl Offer Letter Generator - Taskpane Integration Layer
 *
 * Strategy: Replace placeholders in the CURRENT document. Track previously
 * inserted values so we can find-and-replace them again when the form changes.
 * This allows live updating without destroying the template — the user should
 * first make a copy of the template (SharePoint "Make a copy"), then use the add-in.
 *
 * The add-in shows a one-time reminder to make a copy before generating.
 */

// Track the last values we inserted so we can find and replace them on update
let _lastInserted = {};
let _hasGenerated = false;
let _copyConfirmed = false;
let _liveUpdateTimer = null;
const LIVE_UPDATE_DELAY = 800;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log('Handl Offer Letter Generator ready');
    window.initFormState();
    window.initializeAddIn();
    setupLiveUpdate();
  }
});

/**
 * Attach debounced live update to all form fields.
 * Only fires after first Generate, so the user clicks once to confirm,
 * then subsequent edits auto-update.
 */
function setupLiveUpdate() {
  const form = document.getElementById('offerForm');
  if (!form) return;
  form.addEventListener('input', debouncedLiveUpdate);
  form.addEventListener('change', debouncedLiveUpdate);
}

function debouncedLiveUpdate() {
  // Only live-update after the user has clicked Generate at least once
  if (!_hasGenerated) return;
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
    bonusPctA: raw.f_bonus_pct_a || '0',
    bonusPctB: raw.f_bonus_pct_b || '0',
    bonusDollarA: document.getElementById('f_bonus_dollar_a')?.value || '$0',
    bonusDollarB: document.getElementById('f_bonus_dollar_b')?.value || '$0',
    exempt: raw.f_exempt === 'exempt' ? 'Exempt' : 'Non-Exempt',
    sharesNum: document.getElementById('f_shares_num')?.value || '0',
    sharesPct: raw.f_shares_pct || '0',
    expirationDate: window.formatDate(raw.f_expiration)
  };
}

/**
 * The placeholder map: keys are placeholder IDs, values are the EXACT bracket
 * text found in the template XML. These were verified by parsing the .docx.
 *
 * IMPORTANT: The template has SEPARATE placeholders for each bonus field,
 * not combined ranges. e.g. [BONUS A %] and [BONUS B %] are separate.
 */
// Verified against actual template XML — these are the EXACT placeholders.
// [TITLE] appears 2x (both get same replacement). [EXEMPT] appears 2x but
// needs DIFFERENT handling — handled separately below.
// There is NO [$ SHARES] placeholder in the template.
const PLACEHOLDERS = {
  name:            '[NAME]',
  title:           '[TITLE]',
  startDate:       '[START DATE]',
  supervisor:      '[SUPERVISOR]',
  salary:          '[SALARY]',
  bonusPctA:       '[BONUS A %]',
  bonusPctB:       '[BONUS B %]',
  bonusDollarA:    '[BONUS A $]',
  bonusDollarB:    '[BONUS B $]',
  sharesNum:       '[# OF SHARES]',
  sharesPct:       '[SHARES %]',
  expirationDate:  '[EXPIRATION DATE]'
};

// EXEMPT is handled specially because it appears twice with different meanings:
// 1) "classified as [EXEMPT]" → replace with "Exempt" or "Non-Exempt"
// 2) "you will [EXEMPT] be eligible" → replace with "not" (if Exempt) or "" (if Non-Exempt)
// We handle this by searching for the contextual phrase instead of just [EXEMPT].

/**
 * Generate Letter — replaces placeholders (or previously inserted values) with form data.
 * First click: shows copy reminder. Second click (or after confirm): does replacements.
 */
window.updateDocument = async function() {
  try {
    // Validate required fields
    if (!window.checkFormStatus()) {
      showStatus('Please fill in all required fields.', 'error');
      return;
    }

    // On first click, remind user to make a copy
    if (!_copyConfirmed && !_hasGenerated) {
      _copyConfirmed = true;
      showStatus('Important: Make sure you are working on a COPY of the template, not the original. Click "Generate Letter" again to proceed.', 'info');
      return;
    }

    const formData = buildFormData();
    showStatus('Updating offer letter...', 'info');

    await Word.run(async (context) => {
      const body = context.document.body;

      // For each field, determine what to search for:
      // - First time: search for the original [PLACEHOLDER] bracket text
      // - Subsequent times: search for the previously inserted value
      for (const [key, originalPlaceholder] of Object.entries(PLACEHOLDERS)) {
        const newValue = formData[key];
        if (!newValue && newValue !== '') continue;

        const searchFor = _hasGenerated ? (_lastInserted[key] || originalPlaceholder) : originalPlaceholder;

        // Skip if the value hasn't changed
        if (_hasGenerated && searchFor === newValue) continue;

        const searchResults = body.search(searchFor, { matchCase: true, matchWholeWord: false });
        searchResults.load('items');
        await context.sync();

        for (let i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].insertText(newValue, Word.InsertLocation.replace);
        }

        // Track what we inserted so we can find it next time
        _lastInserted[key] = newValue;
      }

      await context.sync();

      // Handle EXEMPT — two different contexts:
      // 1) "classified as [EXEMPT]" → "Exempt" or "Non-Exempt"
      // 2) "will [EXEMPT] be eligible" → "not" (if Exempt) or "" (if Non-Exempt)
      if (!_hasGenerated) {
        // First generate: search for "classified as [EXEMPT]" context
        // Just replace all [EXEMPT] occurrences — first one gets the label,
        // but since body.search replaces ALL matches, we handle it differently.
        // Replace the contextual phrase "will [EXEMPT] be eligible" FIRST (more specific)
        try {
          const exemptEligible = formData.exempt === 'Exempt' ? 'will not be eligible' : 'will be eligible';
          const willExemptResults = body.search('will [EXEMPT] be eligible', { matchCase: false, matchWholeWord: false });
          willExemptResults.load('items');
          await context.sync();
          for (let i = 0; i < willExemptResults.items.length; i++) {
            willExemptResults.items[i].insertText(exemptEligible, Word.InsertLocation.replace);
          }
          _lastInserted._exemptEligible = exemptEligible;
          await context.sync();
        } catch (e) {
          console.warn('Exempt eligibility replacement:', e);
        }

        // Now replace the remaining [EXEMPT] (the classification label)
        try {
          const classifiedResults = body.search('[EXEMPT]', { matchCase: true, matchWholeWord: false });
          classifiedResults.load('items');
          await context.sync();
          for (let i = 0; i < classifiedResults.items.length; i++) {
            classifiedResults.items[i].insertText(formData.exempt, Word.InsertLocation.replace);
          }
          _lastInserted._exemptLabel = formData.exempt;
          await context.sync();
        } catch (e) {
          console.warn('Exempt label replacement:', e);
        }
      } else {
        // Subsequent updates: search for previously inserted values
        try {
          if (_lastInserted._exemptEligible) {
            const newEligible = formData.exempt === 'Exempt' ? 'will not be eligible' : 'will be eligible';
            if (newEligible !== _lastInserted._exemptEligible) {
              const eligResults = body.search(_lastInserted._exemptEligible, { matchCase: false, matchWholeWord: false });
              eligResults.load('items');
              await context.sync();
              for (let i = 0; i < eligResults.items.length; i++) {
                eligResults.items[i].insertText(newEligible, Word.InsertLocation.replace);
              }
              _lastInserted._exemptEligible = newEligible;
            }
          }
          if (_lastInserted._exemptLabel && _lastInserted._exemptLabel !== formData.exempt) {
            const labelResults = body.search(_lastInserted._exemptLabel, { matchCase: true, matchWholeWord: true });
            labelResults.load('items');
            await context.sync();
            for (let i = 0; i < labelResults.items.length; i++) {
              labelResults.items[i].insertText(formData.exempt, Word.InsertLocation.replace);
            }
            _lastInserted._exemptLabel = formData.exempt;
          }
          await context.sync();
        } catch (e) {
          console.warn('Exempt update:', e);
        }
      }

      // Handle bonus paragraph deletion (only on first generate when toggled off)
      if (!formData.bonusEnabled && !_hasGenerated) {
        try {
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
          await context.sync();
        } catch (e) {
          console.warn('Bonus paragraph deletion:', e);
        }
      }
    });

    _hasGenerated = true;
    showStatus('Offer letter updated! Changes will auto-sync as you edit the form.', 'success');
  } catch (error) {
    console.error('Error generating offer letter:', error);
    showStatus('Error: ' + error.message, 'error');
  }
};

/**
 * Save As: Download the current document with the correct filename
 */
window.saveDocument = async function() {
  try {
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';
    const filename = `Handl_Offer_Letter_${name.replace(/\s+/g, '_')}.docx`;

    showStatus('Preparing download...', 'info');

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
                showStatus('Downloaded as ' + filename, 'success');
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
    console.error('Error downloading:', error);
    const raw = window.getFormData();
    const name = raw.f_name || 'Unknown';
    const filename = `Handl_Offer_Letter_${name.replace(/\s+/g, '_')}.docx`;
    showStatus('To download: File > Save As > Download a Copy. Rename to "' + filename + '"', 'info');
  }
};
