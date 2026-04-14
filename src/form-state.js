/**
 * Form State Management for Offer Letter Generator
 * Handles form initialization, state tracking, validation, and auto-calculations
 */
console.log('[form-state.js] v3.0 loaded');

// Constants
const TOTAL_SHARES = 21231806;
const SHARE_PRICE = 2.21;

// Form state tracker
let formState = {
  f_name: '',
  f_title: '',
  f_start_date: '',
  f_exempt: 'exempt',
  f_supervisor: '',
  f_salary: '',
  bonusToggle: true,
  f_bonus_pct_a: '',
  f_bonus_pct_b: '',
  f_bonus_dollar_a: '',
  f_bonus_dollar_b: '',
  f_shares_pct: '',
  f_shares_num: '',
  f_shares_val: '',
  f_expiration: '',
};

// Required fields for form validation
const REQUIRED_FIELDS = [
  'f_name',
  'f_title',
  'f_start_date',
  'f_supervisor',
  'f_salary',
  'f_expiration',
];

/**
 * Initialize form state and attach event listeners
 */
function initFormState() {
  const form = document.getElementById('offerForm');
  if (!form) {
    console.error('Form element not found');
    return;
  }

  // Salary input - triggers bonus calculations
  const salaryInput = document.getElementById('f_salary');
  if (salaryInput) {
    salaryInput.addEventListener('input', (e) => {
      formState.f_salary = e.target.value;
      updateCalculatedFields();
      checkFormStatus();
    });
  }

  // Bonus toggle
  const bonusToggle = document.getElementById('bonusToggle');
  if (bonusToggle) {
    bonusToggle.addEventListener('change', (e) => {
      formState.bonusToggle = e.target.checked;
      toggleBonusFields(e.target.checked);
      updateCalculatedFields();
      checkFormStatus();
    });
  }

  // Bonus percentage inputs
  const bonusPctA = document.getElementById('f_bonus_pct_a');
  if (bonusPctA) {
    bonusPctA.addEventListener('input', (e) => {
      formState.f_bonus_pct_a = e.target.value;
      updateCalculatedFields();
      checkFormStatus();
    });
  }

  const bonusPctB = document.getElementById('f_bonus_pct_b');
  if (bonusPctB) {
    bonusPctB.addEventListener('input', (e) => {
      formState.f_bonus_pct_b = e.target.value;
      updateCalculatedFields();
      checkFormStatus();
    });
  }

  // Shares percentage input - triggers equity calculations
  const sharesPct = document.getElementById('f_shares_pct');
  if (sharesPct) {
    sharesPct.addEventListener('input', (e) => {
      formState.f_shares_pct = e.target.value;
      updateCalculatedFields();
      checkFormStatus();
    });
  }

  // Required text/date inputs
  const requiredInputs = [
    'f_name',
    'f_title',
    'f_start_date',
    'f_supervisor',
    'f_expiration',
  ];

  requiredInputs.forEach((fieldId) => {
    const input = document.getElementById(fieldId);
    if (input) {
      input.addEventListener('input', (e) => {
        formState[fieldId] = e.target.value;
        checkFormStatus();
      });
      input.addEventListener('change', (e) => {
        formState[fieldId] = e.target.value;
        checkFormStatus();
      });
    }
  });

  // FLSA Status select
  const exemptSelect = document.getElementById('f_exempt');
  if (exemptSelect) {
    exemptSelect.addEventListener('change', (e) => {
      formState.f_exempt = e.target.value;
    });
  }

  // Action buttons
  const generateBtn = document.getElementById('generateLetterBtn');
  if (generateBtn) {
    generateBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (typeof window.generateAndDownload === 'function') {
        window.generateAndDownload();
      } else if (typeof window.updateDocument === 'function') {
        window.updateDocument();
      } else {
        console.warn('No generate function available');
      }
    });
  }

  // Initialize bonus fields visibility
  toggleBonusFields(formState.bonusToggle);

  // Initial form status check
  checkFormStatus();

  // NOTE: Pre-caching on focusin was removed — getFileAsync hangs from focusin
  // in Word Online. Template bytes are now cached in generateAndDownload (button click).

  // Live preview: trigger on focusout (when leaving a field) AND input (while typing).
  // focusout fires once per field (clean), input fires per keystroke (responsive).
  // The 100ms debounce + previewRunning lock prevent excessive Word.run calls.
  if (form) {
    form.addEventListener('focusout', function (e) {
      if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT')) {
        console.log('[focusout] field=' + e.target.id);
        if (typeof window.schedulePreviewRefresh === 'function') {
          window.schedulePreviewRefresh();
        } else {
          console.warn('[focusout] schedulePreviewRefresh not defined!');
        }
      }
    });
    form.addEventListener('input', function (e) {
      if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT')) {
        if (typeof window.schedulePreviewRefresh === 'function') {
          window.schedulePreviewRefresh();
        }
      }
    });
    form.addEventListener('change', function (e) {
      if (e.target && (e.target.tagName === 'SELECT' || e.target.type === 'checkbox')) {
        console.log('[change] field=' + e.target.id);
        if (typeof window.schedulePreviewRefresh === 'function') {
          window.schedulePreviewRefresh();
        }
      }
    });
    console.log('[initFormState] Preview event listeners attached (focusout + input + change)');
  }
}

/**
 * Get all current form values
 * @returns {Object} Form data object
 */
function getFormData() {
  return {
    f_name: document.getElementById('f_name')?.value || '',
    f_title: document.getElementById('f_title')?.value || '',
    f_start_date: document.getElementById('f_start_date')?.value || '',
    f_exempt: document.getElementById('f_exempt')?.value || 'exempt',
    f_supervisor: document.getElementById('f_supervisor')?.value || '',
    f_salary: document.getElementById('f_salary')?.value || '',
    bonusToggle: document.getElementById('bonusToggle')?.checked || false,
    f_bonus_pct_a: document.getElementById('f_bonus_pct_a')?.value || '',
    f_bonus_pct_b: document.getElementById('f_bonus_pct_b')?.value || '',
    f_bonus_dollar_a: document.getElementById('f_bonus_dollar_a')?.value || '',
    f_bonus_dollar_b: document.getElementById('f_bonus_dollar_b')?.value || '',
    f_shares_pct: document.getElementById('f_shares_pct')?.value || '',
    f_shares_num: document.getElementById('f_shares_num')?.value || '',
    f_shares_val: document.getElementById('f_shares_val')?.value || '',
    f_expiration: document.getElementById('f_expiration')?.value || '',
  };
}

/**
 * Calculate bonus amounts based on salary and percentages
 * @param {number|string} salary Annual salary
 * @param {number|string} pctA Minimum bonus percentage
 * @param {number|string} pctB Maximum bonus percentage
 * @returns {Object} { minDollar, maxDollar } rounded to nearest dollar
 */
function calculateBonus(salary, pctA, pctB) {
  const salaryNum = parseFloat(salary) || 0;
  const pctANum = parseFloat(pctA) || 0;
  const pctBNum = parseFloat(pctB) || 0;

  if (salaryNum <= 0) {
    return { minDollar: 0, maxDollar: 0 };
  }

  const minDollar = Math.round((salaryNum * pctANum) / 100);
  const maxDollar = Math.round((salaryNum * pctBNum) / 100);

  return { minDollar, maxDollar };
}

/**
 * Calculate equity amounts based on shares percentage
 * @param {number|string} sharesPct Equity percentage
 * @returns {Object} { numShares, sharesValue } with calculated values
 */
function calculateEquity(sharesPct) {
  const pctNum = parseFloat(sharesPct) || 0;

  if (pctNum <= 0) {
    return { numShares: 0, sharesValue: 0 };
  }

  // Calculate number of shares: round(pct / 100 × 21,231,806)
  const numShares = Math.round((pctNum / 100) * TOTAL_SHARES);

  // Calculate shares value: round(numShares × 2.21)
  const sharesValue = Math.round(numShares * SHARE_PRICE);

  return { numShares, sharesValue };
}

/**
 * Update all calculated fields when inputs change
 */
function updateCalculatedFields() {
  // Update bonus calculations if toggle is on
  if (formState.bonusToggle) {
    const salary = document.getElementById('f_salary')?.value || '';
    const pctA = document.getElementById('f_bonus_pct_a')?.value || '';
    const pctB = document.getElementById('f_bonus_pct_b')?.value || '';

    if (salary && (pctA || pctB)) {
      const { minDollar, maxDollar } = calculateBonus(salary, pctA, pctB);

      const minDollarInput = document.getElementById('f_bonus_dollar_a');
      if (minDollarInput) {
        minDollarInput.value = formatCurrency(minDollar);
      }

      const maxDollarInput = document.getElementById('f_bonus_dollar_b');
      if (maxDollarInput) {
        maxDollarInput.value = formatCurrency(maxDollar);
      }

      formState.f_bonus_dollar_a = formatCurrency(minDollar);
      formState.f_bonus_dollar_b = formatCurrency(maxDollar);
    } else {
      const minDollarInput = document.getElementById('f_bonus_dollar_a');
      if (minDollarInput) {
        minDollarInput.value = '';
      }

      const maxDollarInput = document.getElementById('f_bonus_dollar_b');
      if (maxDollarInput) {
        maxDollarInput.value = '';
      }

      formState.f_bonus_dollar_a = '';
      formState.f_bonus_dollar_b = '';
    }
  }

  // Update equity calculations
  const sharesPct = document.getElementById('f_shares_pct')?.value || '';

  if (sharesPct) {
    const { numShares, sharesValue } = calculateEquity(sharesPct);

    const sharesNumInput = document.getElementById('f_shares_num');
    if (sharesNumInput) {
      sharesNumInput.value =
        numShares > 0 ? numShares.toLocaleString() : '';
    }

    const sharesValInput = document.getElementById('f_shares_val');
    if (sharesValInput) {
      sharesValInput.value = sharesValue > 0 ? formatCurrency(sharesValue) : '';
    }

    formState.f_shares_num = numShares.toString();
    formState.f_shares_val = formatCurrency(sharesValue);
  } else {
    const sharesNumInput = document.getElementById('f_shares_num');
    if (sharesNumInput) {
      sharesNumInput.value = '';
    }

    const sharesValInput = document.getElementById('f_shares_val');
    if (sharesValInput) {
      sharesValInput.value = '';
    }

    formState.f_shares_num = '';
    formState.f_shares_val = '';
  }
}

/**
 * Show or hide bonus fields section
 * @param {boolean} show Whether to show the bonus fields
 */
function toggleBonusFields(show) {
  const bonusFieldsContainer = document.getElementById('bonusFields');
  if (!bonusFieldsContainer) return;

  if (show) {
    bonusFieldsContainer.classList.remove('hidden');
  } else {
    bonusFieldsContainer.classList.add('hidden');
    // Clear bonus values when toggled off
    document.getElementById('f_bonus_pct_a').value = '';
    document.getElementById('f_bonus_pct_b').value = '';
    document.getElementById('f_bonus_dollar_a').value = '';
    document.getElementById('f_bonus_dollar_b').value = '';

    formState.f_bonus_pct_a = '';
    formState.f_bonus_pct_b = '';
    formState.f_bonus_dollar_a = '';
    formState.f_bonus_dollar_b = '';
  }
}

/**
 * Check if all required fields are filled and update status pill
 * @returns {boolean} True if form is complete, false otherwise
 */
function checkFormStatus() {
  const allFilled = REQUIRED_FIELDS.every((fieldId) => {
    const input = document.getElementById(fieldId);
    return input && input.value && input.value.trim() !== '';
  });

  const statusPill = document.getElementById('statusPill');
  if (statusPill) {
    if (allFilled) {
      statusPill.textContent = '';
      statusPill.classList.remove('status-draft');
      statusPill.classList.add('status-ready');
      statusPill.innerHTML =
        '<span class="status-dot"></span>Ready';
    } else {
      statusPill.textContent = '';
      statusPill.classList.remove('status-ready');
      statusPill.classList.add('status-draft');
      statusPill.innerHTML =
        '<span class="status-dot"></span>Draft';
    }
  }

  return allFilled;
}

/**
 * Format number as currency
 * @param {number} num Number to format
 * @returns {string} Formatted currency string (e.g., "$1,234")
 */
function formatCurrency(num) {
  const numVal = parseInt(num, 10) || 0;
  return '$' + numVal.toLocaleString();
}

/**
 * Format date string to readable format
 * @param {string} dateStr Date string in YYYY-MM-DD format
 * @returns {string} Formatted date string (e.g., "April 9, 2026")
 */
function formatDate(dateStr) {
  if (!dateStr) return '';

  try {
    const date = new Date(dateStr + 'T00:00:00');
    if (isNaN(date.getTime())) return '';

    var formatted = date.toLocaleDateString('en-US', {
      month: 'long',
      day: 'numeric',
    });
    // Always use 4-digit year (toLocaleDateString can truncate leading zeros)
    var year = date.getFullYear().toString().padStart(4, '0');
    return formatted + ', ' + year;
  } catch (e) {
    return '';
  }
}

/**
 * Reset the form to its initial blank state
 */
function resetForm() {
  const form = document.getElementById('offerForm');
  if (form) form.reset();

  // Reset internal state tracking
  formState = {
    f_name: '', f_title: '', f_start_date: '', f_exempt: 'exempt',
    f_supervisor: '', f_salary: '', bonusToggle: true,
    f_bonus_pct_a: '', f_bonus_pct_b: '', f_bonus_dollar_a: '',
    f_bonus_dollar_b: '', f_shares_pct: '', f_shares_num: '',
    f_shares_val: '', f_expiration: ''
  };

  // Clear calculated field displays (form.reset doesn't clear readonly fields set via JS)
  ['f_bonus_dollar_a', 'f_bonus_dollar_b', 'f_shares_num', 'f_shares_val'].forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.value = '';
  });

  // Restore bonus section visibility and status pill
  toggleBonusFields(true);
  checkFormStatus();
}

// Export functions to window object for cross-file access
window.resetForm = resetForm;
window.initFormState = initFormState;
window.getFormData = getFormData;
window.calculateBonus = calculateBonus;
window.calculateEquity = calculateEquity;
window.updateCalculatedFields = updateCalculatedFields;
window.toggleBonusFields = toggleBonusFields;
window.checkFormStatus = checkFormStatus;
window.formatCurrency = formatCurrency;
window.formatDate = formatDate;
