/**
 * Handl Offer Letter Generator - Taskpane Integration Layer
 *
 * STRATEGY: Capture → Replace in Memory → Download
 * 1. User fills in form on the taskpane
 * 2. Clicks "Generate & Download"
 * 3. Add-in reads the template document bytes via getFileAsync
 * 4. Uses JSZip to open the .docx in memory
 * 5. Replaces all placeholders in word/document.xml
 * 6. Downloads the modified .docx as Handl_Offer_Letter_[Name].docx
 *
 * The template is NEVER modified. All work happens in memory.
 */
console.log('[taskpane.js] v2.2 loaded');

// Cache template bytes so we only call getFileAsync once.
// This fixes the "Error reading template" bug on 2nd+ generation —
// Word Online doesn't reliably support rapid consecutive getFileAsync calls.
let cachedTemplateBytes = null;

// Cache the original body OOXML so we can restore it wholesale on revert.
// This avoids the race condition of field-by-field replacement during form reset.
let cachedBodyOoxml = null;

// ===== LIVE PREVIEW SYSTEM =====
// Tracks what's currently displayed in the document for each placeholder.
// On field change: search for current text, replace with form value.
// On generate+reset: restore cached original OOXML wholesale.

const SIMPLE_PLACEHOLDERS = [
  '[NAME]', '[TITLE]', '[START DATE]', '[SUPERVISOR]', '[SALARY]',
  '[BONUS A %]', '[BONUS B %]', '[BONUS A $]', '[BONUS B $]',
  '[# OF SHARES]', '[SHARES %]', '[EXPIRATION DATE]'
];

// Current displayed text for each tracked slot
const previewState = {};
for (const p of SIMPLE_PLACEHOLDERS) {
  previewState[p] = p;
}
// EXEMPT appears twice with different contexts — track separately
previewState['__EXEMPT_CLASS__'] = '[EXEMPT]';
previewState['__EXEMPT_ELIG__'] = 'will [EXEMPT] be eligible';

let previewRefreshTimer = null;
let previewRunning = false; // Lock to prevent concurrent Word.run calls
let revertInProgress = false; // Hard block: suppresses ALL preview scheduling during revert
let ooxmlCacheInProgress = false; // Blocks preview while OOXML caching Word.run is active

/**
 * Compute what each placeholder slot should currently display, based on form state.
 * Empty fields fall back to the original placeholder text (no change in doc).
 */
function getDesiredPreview() {
  const raw = window.getFormData();
  const d = {};
  d['[NAME]'] = raw.f_name || '[NAME]';
  d['[TITLE]'] = raw.f_title || '[TITLE]';
  d['[START DATE]'] = raw.f_start_date ? window.formatDate(raw.f_start_date) : '[START DATE]';
  d['[SUPERVISOR]'] = raw.f_supervisor || '[SUPERVISOR]';
  d['[SALARY]'] = raw.f_salary ? window.formatCurrency(parseFloat(raw.f_salary) || 0) : '[SALARY]';
  d['[BONUS A %]'] = raw.f_bonus_pct_a || '[BONUS A %]';
  d['[BONUS B %]'] = raw.f_bonus_pct_b || '[BONUS B %]';
  d['[BONUS A $]'] = document.getElementById('f_bonus_dollar_a')?.value || '[BONUS A $]';
  d['[BONUS B $]'] = document.getElementById('f_bonus_dollar_b')?.value || '[BONUS B $]';
  d['[# OF SHARES]'] = document.getElementById('f_shares_num')?.value || '[# OF SHARES]';
  d['[SHARES %]'] = raw.f_shares_pct || '[SHARES %]';
  d['[EXPIRATION DATE]'] = raw.f_expiration ? window.formatDate(raw.f_expiration) : '[EXPIRATION DATE]';

  const isExempt = raw.f_exempt === 'exempt';
  d['__EXEMPT_CLASS__'] = isExempt ? 'Exempt' : 'Non-Exempt';
  d['__EXEMPT_ELIG__'] = isExempt ? 'will not be eligible' : 'will be eligible';
  return d;
}

/**
 * Replace text in the document using OOXML get/set.
 * Avoids the Word search API entirely (which hangs on bracket characters).
 * Gets the body OOXML, does string replacement in JS, sets it back.
 * @param {Array<{from: string, to: string}>} changes
 */
async function replaceInDocument(changes) {
  if (changes.length === 0) return;

  // Save current focus so we can restore it after insertOoxml steals it
  var savedEl = document.activeElement;
  var savedStart = null;
  var savedEnd = null;
  if (savedEl && typeof savedEl.selectionStart === 'number') {
    savedStart = savedEl.selectionStart;
    savedEnd = savedEl.selectionEnd;
  }

  try {
    await Word.run(async (context) => {
      // Get the body OOXML (contains all text + formatting)
      var ooxmlResult = context.document.body.getOoxml();
      await context.sync();

      // Do all replacements on the XML string
      var xml = ooxmlResult.value;
      for (var i = 0; i < changes.length; i++) {
        var fromXml = escapeXml(changes[i].from);
        var toXml = escapeXml(changes[i].to);
        var before = xml;
        xml = xml.split(fromXml).join(toXml);
        var replaced = (before !== xml);
        console.log((replaced ? '✓' : '✗') + ' "' + changes[i].from + '" → "' + changes[i].to + '"');
      }

      // Set the modified OOXML back
      context.document.body.insertOoxml(xml, "Replace");
      await context.sync();
    });
  } catch (e) {
    console.error('replaceInDocument error:', e);
  }

  // Restore focus to the form field that was active before insertOoxml stole it
  if (savedEl && savedEl.closest && savedEl.closest('#offerForm')) {
    try {
      savedEl.focus();
      if (savedStart !== null && typeof savedEl.setSelectionRange === 'function') {
        savedEl.setSelectionRange(savedStart, savedEnd);
      }
    } catch (e) { /* ignore */ }
  }
}

/**
 * Refresh the live preview — compare desired values to current state, update changed ones.
 * Uses a single batched Word.run call for all changes.
 */
async function refreshPreview() {
  const desired = getDesiredPreview();
  const changes = [];

  // EXEMPT eligibility FIRST (more specific phrase match)
  if (previewState['__EXEMPT_ELIG__'] !== desired['__EXEMPT_ELIG__']) {
    changes.push({ from: previewState['__EXEMPT_ELIG__'], to: desired['__EXEMPT_ELIG__'], key: '__EXEMPT_ELIG__' });
  }
  // EXEMPT classification
  if (previewState['__EXEMPT_CLASS__'] !== desired['__EXEMPT_CLASS__']) {
    changes.push({ from: previewState['__EXEMPT_CLASS__'], to: desired['__EXEMPT_CLASS__'], key: '__EXEMPT_CLASS__' });
  }
  // All simple placeholders
  for (const p of SIMPLE_PLACEHOLDERS) {
    if (previewState[p] !== desired[p]) {
      changes.push({ from: previewState[p], to: desired[p], key: p });
    }
  }

  if (changes.length === 0) return;

  // Lazy-cache body OOXML on first preview (before any modifications).
  // Runs under the previewRunning lock so no other Word.run can overlap.
  if (!cachedBodyOoxml) {
    console.log('Caching body OOXML before first preview...');
    try {
      await Word.run(async (context) => {
        var ooxmlResult = context.document.body.getOoxml();
        await context.sync();
        cachedBodyOoxml = ooxmlResult.value;
        console.log('Body OOXML cached (' + cachedBodyOoxml.length + ' chars)');
      });
    } catch (e) {
      console.warn('OOXML cache failed (field-by-field revert fallback available):', e);
    }
  }

  console.log('Preview: updating ' + changes.length + ' field(s)');

  await replaceInDocument(changes);

  // Update state after successful batch
  for (var i = 0; i < changes.length; i++) {
    previewState[changes[i].key] = changes[i].to;
  }
}

/**
 * Schedule a preview refresh with short debounce (100ms).
 * Triggered on field blur (not every keystroke), so the debounce is just
 * to coalesce rapid-fire blur/focus events when tabbing between fields.
 */
window.schedulePreviewRefresh = function () {
  if (revertInProgress) {
    console.log('Preview suppressed — revert in progress');
    return;
  }
  if (previewRefreshTimer) clearTimeout(previewRefreshTimer);
  previewRefreshTimer = setTimeout(async function () {
    if (revertInProgress) return;
    if (previewRunning || ooxmlCacheInProgress) {
      console.log('Preview busy, re-scheduling...');
      previewRefreshTimer = setTimeout(function () {
        window.schedulePreviewRefresh();
      }, 1000);
      return;
    }
    previewRunning = true;
    try {
      console.log('Preview refresh starting...');
      await refreshPreview();
      console.log('Preview refresh complete');
    } catch (e) {
      console.error('Preview refresh error:', e);
    } finally {
      previewRunning = false;
    }
  }, 300);
};

/**
 * Revert document to original template state by restoring cached body OOXML wholesale.
 * This avoids the race condition of field-by-field replacement during form reset.
 * Falls back to field-by-field if cached OOXML is not available.
 */
async function revertDocument() {
  if (cachedBodyOoxml) {
    // WHOLESALE RESTORE — single insertOoxml call, no field-by-field matching needed
    console.log('Reverting document via cached OOXML restore...');
    try {
      await Word.run(async (context) => {
        context.document.body.insertOoxml(cachedBodyOoxml, "Replace");
        await context.sync();
      });
      console.log('Document restored to original template');
    } catch (e) {
      console.error('Wholesale OOXML restore failed:', e);
    }
  } else {
    // Fallback: field-by-field replacement (less reliable)
    console.warn('No cached OOXML — falling back to field-by-field revert');
    const changes = [];
    for (const p of SIMPLE_PLACEHOLDERS) {
      if (previewState[p] !== p) {
        changes.push({ from: previewState[p], to: p, key: p });
      }
    }
    if (previewState['__EXEMPT_CLASS__'] !== '[EXEMPT]') {
      changes.push({ from: previewState['__EXEMPT_CLASS__'], to: '[EXEMPT]', key: '__EXEMPT_CLASS__' });
    }
    if (previewState['__EXEMPT_ELIG__'] !== 'will [EXEMPT] be eligible') {
      changes.push({ from: previewState['__EXEMPT_ELIG__'], to: 'will [EXEMPT] be eligible', key: '__EXEMPT_ELIG__' });
    }
    if (changes.length > 0) {
      console.log('Reverting ' + changes.length + ' placeholder(s) field-by-field');
      await replaceInDocument(changes);
    }
  }

  // Reset previewState to original placeholders regardless of method
  for (const p of SIMPLE_PLACEHOLDERS) {
    previewState[p] = p;
  }
  previewState['__EXEMPT_CLASS__'] = '[EXEMPT]';
  previewState['__EXEMPT_ELIG__'] = 'will [EXEMPT] be eligible';
}

/**
 * Pre-cache template bytes is NO LONGER triggered on focusin.
 * getFileAsync hangs from focusin in Word Online (not a valid user gesture).
 * Template bytes are now cached on-demand in generateAndDownload() (button click).
 * Body OOXML is cached lazily in refreshPreview() under the previewRunning lock.
 * This function is kept as a no-op for backward compatibility.
 */
window.preCacheTemplateBytes = async function () {
  console.log('preCacheTemplateBytes called (no-op — caching happens on demand)');
};

Office.onReady(function (info) {
  console.log('[Office.onReady] host=' + info.host + ', platform=' + info.platform);
  if (info.host === Office.HostType.Word) {
    console.log('Handl Offer Letter Generator ready — initializing form...');
    window.initFormState();
    window.initializeAddIn();
    console.log('Form initialized, event listeners attached');
  } else {
    console.warn('Not running in Word — host is: ' + info.host);
  }
});

/**
 * Build form data object with formatted values for placeholder replacement
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
 * Read the current document as a Uint8Array via Office.js getFileAsync
 */
function getDocumentBytes() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 262144 },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error ? result.error.message : 'Failed to get file'));
          return;
        }
        const file = result.value;
        const sliceCount = file.sliceCount;
        const sliceData = [];

        function readSlice(index) {
          file.getSliceAsync(index, function (sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
              sliceData.push(sliceResult.value.data);
              if (index + 1 < sliceCount) {
                readSlice(index + 1);
              } else {
                file.closeAsync();
                // Combine all slices into one Uint8Array
                let totalLength = 0;
                const arrays = sliceData.map((slice) => {
                  const arr = new Uint8Array(slice);
                  totalLength += arr.length;
                  return arr;
                });
                const combined = new Uint8Array(totalLength);
                let offset = 0;
                for (const arr of arrays) {
                  combined.set(arr, offset);
                  offset += arr.length;
                }
                resolve(combined);
              }
            } else {
              file.closeAsync();
              reject(new Error('Failed to read file slice'));
            }
          });
        }
        readSlice(0);
      }
    );
  });
}

/**
 * Trigger browser download of a byte array as .docx
 */
function downloadDocx(bytes, filename) {
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
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

/**
 * Escape special XML characters in replacement values
 */
function escapeXml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Replace placeholders in the document XML string.
 * Handles placeholders that may be split across multiple XML runs
 * by first joining adjacent <w:t> text nodes, doing replacements,
 * then returning the modified XML.
 */
function replacePlaceholders(xml, formData) {
  // Define all placeholder → value mappings
  const replacements = [
    { find: '[NAME]', replace: formData.name },
    { find: '[TITLE]', replace: formData.title },
    { find: '[START DATE]', replace: formData.startDate },
    { find: '[SUPERVISOR]', replace: formData.supervisor },
    { find: '[SALARY]', replace: formData.salary },
    { find: '[BONUS A %]', replace: formData.bonusPctA },
    { find: '[BONUS B %]', replace: formData.bonusPctB },
    { find: '[BONUS A $]', replace: formData.bonusDollarA },
    { find: '[BONUS B $]', replace: formData.bonusDollarB },
    { find: '[# OF SHARES]', replace: formData.sharesNum },
    { find: '[SHARES %]', replace: formData.sharesPct },
    { find: '[EXPIRATION DATE]', replace: formData.expirationDate },
  ];

  let result = xml;

  // Handle EXEMPT contextual replacement first (more specific match)
  // "will [EXEMPT] be eligible" → "will not be eligible" (Exempt) or "will be eligible" (Non-Exempt)
  const exemptPhrase =
    formData.exempt === 'Exempt' ? 'will not be eligible' : 'will be eligible';

  // Try direct text replacement first (placeholder in single run)
  result = result.split('will [EXEMPT] be eligible').join(escapeXml(exemptPhrase));

  // Also handle the case where [EXEMPT] is in its own XML run within the phrase
  // Pattern: ...will </w:t></w:r>...<w:r>...<w:t...>[EXEMPT]</w:t></w:r>...<w:r>...<w:t...> be eligible...
  // We handle this with a regex that matches across runs
  const exemptRunRegex =
    /(will\s*<\/w:t><\/w:r>[\s\S]*?<w:r[^>]*>[\s\S]*?<w:t[^>]*>)\[EXEMPT\](<\/w:t><\/w:r>[\s\S]*?<w:r[^>]*>[\s\S]*?<w:t[^>]*>\s*be eligible)/g;
  result = result.replace(exemptRunRegex, function () {
    // Replace the whole matched phrase with the resolved text in the first run
    return escapeXml(exemptPhrase);
  });

  // Replace remaining standalone [EXEMPT] with the classification label
  result = result.split('[EXEMPT]').join(escapeXml(formData.exempt));

  // Replace all standard placeholders
  for (const { find, replace } of replacements) {
    if (replace === undefined || replace === null) continue;
    // Simple string split/join — works when placeholder is in a single <w:t> run
    result = result.split(find).join(escapeXml(replace));
  }

  // Handle bonus paragraph deletion if bonus is toggled off
  if (!formData.bonusEnabled) {
    result = deleteBonusParagraph(result);
  }

  return result;
}

/**
 * Remove the paragraph containing "Discretionary, performance-based bonus" from XML
 */
function deleteBonusParagraph(xml) {
  // Find <w:p> elements that contain the bonus text and remove them
  const bonusText = 'Discretionary, performance-based bonus';
  const paragraphRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;

  return xml.replace(paragraphRegex, function (match) {
    if (match.indexOf(bonusText) !== -1) {
      return ''; // Remove this paragraph entirely
    }
    return match;
  });
}

/**
 * MAIN ACTION: Generate & Download
 *
 * 1. Read the template document bytes (getFileAsync) — document is NOT modified
 * 2. Use JSZip to open the .docx in memory
 * 3. Replace all placeholders in word/document.xml
 * 4. Generate the modified .docx and trigger browser download
 */
window.generateAndDownload = async function () {
  // IMMEDIATELY block all preview scheduling to prevent focusout-triggered
  // Word.run from racing with the revert's Word.run (concurrent = deadlock)
  revertInProgress = true;
  if (previewRefreshTimer) {
    clearTimeout(previewRefreshTimer);
    previewRefreshTimer = null;
  }

  try {
    if (!window.checkFormStatus()) {
      showStatus('Please fill in all required fields.', 'error');
      revertInProgress = false;
      return;
    }

    const formData = buildFormData();
    const filename = `Handl_Offer_Letter_${formData.name.replace(/\s+/g, '_')}.docx`;

    // Disable button while processing
    const btn = document.getElementById('generateLetterBtn');
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Generating & Resetting...';
    }

    // STEP 1: Get template bytes (pre-cached on first form focus)
    let templateBytes;
    try {
      if (cachedTemplateBytes) {
        console.log('Using cached template bytes');
        templateBytes = cachedTemplateBytes;
      } else {
        // Fallback: try reading now (we're in a button click = user gesture)
        console.log('Cache miss — reading template bytes from button click...');
        templateBytes = await getDocumentBytes();
        cachedTemplateBytes = templateBytes;
        console.log('Template bytes cached (' + templateBytes.length + ' bytes)');
      }
    } catch (e) {
      console.error('Failed to read template:', e);
      showStatus('Error reading template. Please try again.', 'error');
      return;
    }

    // STEP 2: Open .docx with JSZip and replace placeholders in memory
    let modifiedBytes;
    try {
      const zip = await JSZip.loadAsync(templateBytes);

      // Get the main document XML
      const docXml = await zip.file('word/document.xml').async('string');

      // Replace all placeholders
      const modifiedXml = replacePlaceholders(docXml, formData);

      // Update the zip with modified XML
      zip.file('word/document.xml', modifiedXml);

      // Also check headers/footers for placeholders (some templates use them)
      const headerFooterFiles = Object.keys(zip.files).filter(
        (name) => name.match(/^word\/(header|footer)\d*\.xml$/)
      );
      for (const hfFile of headerFooterFiles) {
        const hfXml = await zip.file(hfFile).async('string');
        const modifiedHfXml = replacePlaceholders(hfXml, formData);
        zip.file(hfFile, modifiedHfXml);
      }

      // Generate the modified .docx
      modifiedBytes = await zip.generateAsync({ type: 'uint8array' });
    } catch (e) {
      console.error('Failed to process document:', e);
      showStatus('Error processing document. Please try again.', 'error');
      return;
    }

    // STEP 3: Download
    downloadDocx(modifiedBytes, filename);
    showStatus('Downloaded ' + filename, 'success');

    // STEP 4: Ensure preview lock is held for the revert Word.run
    // (revertInProgress was already set at function entry to block focusout races)
    previewRunning = true;

    // STEP 5: Reset form (triggers input/change events, but revertInProgress suppresses them)
    if (typeof window.resetForm === 'function') {
      window.resetForm();
    }

    // STEP 6: Revert document (wholesale OOXML restore), then unblock
    // 10-second timeout safety valve — download already succeeded, revert is best-effort
    try {
      await Promise.race([
        revertDocument(),
        new Promise(function (_, reject) {
          setTimeout(function () { reject(new Error('Revert timed out after 10s')); }, 10000);
        })
      ]);
      console.log('Document reverted to placeholders');
    } catch (e) {
      console.error('Revert failed (download succeeded):', e);
    }
  } catch (error) {
    console.error('Error:', error);
    showStatus('Error: ' + error.message, 'error');
  } finally {
    // Always clear flags and re-enable button, even if revert failed
    previewRunning = false;
    revertInProgress = false;
    const btn = document.getElementById('generateLetterBtn');
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Generate Offer & Reset Form';
    }
  }
};

// Alias both button handlers to the same function
window.updateDocument = window.generateAndDownload;
window.saveDocument = window.generateAndDownload;
