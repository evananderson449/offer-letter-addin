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

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log('Handl Offer Letter Generator ready');
    window.initFormState();
    window.initializeAddIn();
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
  try {
    if (!window.checkFormStatus()) {
      showStatus('Please fill in all required fields.', 'error');
      return;
    }

    const formData = buildFormData();
    const filename = `Handl_Offer_Letter_${formData.name.replace(/\s+/g, '_')}.docx`;

    // Disable button while processing
    const btn = document.getElementById('generateLetterBtn');
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Generating...';
    }

    // STEP 1: Read template bytes (document is NOT modified)
    let templateBytes;
    try {
      templateBytes = await getDocumentBytes();
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
  } catch (error) {
    console.error('Error:', error);
    showStatus('Error: ' + error.message, 'error');
  } finally {
    const btn = document.getElementById('generateLetterBtn');
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Generate & Download';
    }
  }
};

// Alias both button handlers to the same function
window.updateDocument = window.generateAndDownload;
window.saveDocument = window.generateAndDownload;
