// ============================================================================
// DCL REVIEW & SUBMIT - CORRECTED & COMPLETE
// 100% Business Requirements Compliant
// ============================================================================
const CONFIG = {
  POWER_AUTOMATE_FLOW_URL: 'https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b0d4e59cd8d64665b0bdf63aaf5bce3e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gUJP6Qmj-SyKBWSBM0FVSlRrf8qk9MW8oyZYZljF9Q0',
  FLOW_TIMEOUT: 120000,
  REFRESH_INTERVAL: 5000,

  // Document ordering priority (as per requirements)
  DOCUMENT_ORDER: {
    'DCL Template': 1,
    'Loading Plan': 2,
    'Commercial Invoice': 3,
    'Packing List': 4
    // Everything else gets priority 999
  }
};

const STATE = {
  dclNumber: null,
  dclMasterId: null,
  dclMasterData: null,
  dclBrands: [],
  allDocuments: [],
  documentsForMerge: [],
  selectedDocuments: new Set(), // 🆕 ADDED
  mergedPdf: null,
  isSubmitting: false,
  autoRefreshTimer: null,
  loadingPlanOrders: [],
  customerMandatoryDocs: [] // Loaded from customer's cr650_mandatorydocuments field
};

// Customer mandatory doc codes → { label, search terms to match cr650_doc_type directly }
// search: array of keywords matched against cr650_doc_type in cr650_dcl_documents
const MANDATORY_DOC_CONFIG = {
  'PI':                                       { label: 'Proforma Invoice (PI)',               search: ['proforma invoice'] },
  'CI':                                       { label: 'Commercial Invoice (CI)',              search: ['commercial invoice'] },
  'PL':                                       { label: 'Packing List (PL)',                    search: ['packing list'] },
  'DNs':                                      { label: 'Delivery Note (DN)',                   search: ['delivery note'] },
  'System Invoices':                          { label: 'System Invoice (SI)',                  search: ['system invoice'] },
  'COO':                                      { label: 'Certificate of Origin',               search: ['certificate of origin', 'coo'] },
  'Shahada Mansha':                           { label: 'Shahada Mansha',                      search: ['shahada mansha', 'shahada'] },
  'MOFA':                                     { label: 'MOFA Document',                       search: ['mofa'] },
  'Inspection Certificate':                   { label: 'Inspection Certificate',              search: ['inspection certificate'] },
  'Insurance Certificate':                    { label: 'Insurance Certificate',               search: ['insurance certificate'] },
  'ED':                                       { label: 'Customs Export Declaration (ED)',      search: ['customs export', 'export declaration', 'bayan'] },
  'Cust. Exit/Entry Cert.':                   { label: 'Customs Exit / Entry Certificate',    search: ['customs exit', 'entry certificate'] },
  'BL/TCN/Main.':                             { label: 'BL / TCN / Manifest',                 search: ['bl', 'tcn', 'manifest', 'bill of lading'] },
  'AWB Receipt by DHL/Email/WATSAPP/BANK':    { label: 'AWB Receipt',                         search: ['awb'] },
  'Finance Submission For Vendor Freight Invoice': { label: 'Finance Submission (Vendor Freight)', search: ['finance submission', 'invoice submission'] },
  'Vendor Freight Invoice':                   { label: 'Vendor Freight Invoice',              search: ['vendor freight', 'vendor invoice'] },
  'Oracle Freight PO':                        { label: 'Oracle Freight PO',                   search: ['oracle freight', 'oracle po'] },
  'Inspection Invoice':                       { label: 'Inspection Invoice',                  search: ['inspection invoice'] },
  'Oracle Inspection PO':                     { label: 'Oracle Inspection PO',                search: ['oracle inspection'] },
  'Documentation Charges':                    { label: 'Documentation Charges',               search: ['documentation charges'] },
  'Finance Submission':                       { label: 'Finance Submission',                  search: ['finance submission', 'invoice submission'] }
};

/**
 * Checks if a mandatory document is present in the DCL's uploaded documents.
 * Matches search terms against cr650_doc_type in cr650_dcl_documents directly.
 */
function isMandatoryDocPresent(searchTerms) {
  return STATE.allDocuments.some(doc => {
    const docType = (doc.cr650_doc_type || '').toLowerCase();
    if (docType === 'merged pdf' || docType === 'merged dcl package') return false;
    return searchTerms.some(term => docType.includes(term.toLowerCase()));
  });
}


// ============================================================================
// DOCUMENT TYPE MAPPING (System → Checklist)
// 100% accurate mapping to match doc_type to checklist items
// ============================================================================
STATE.checklistMaster = [
  "Proforma Invoice (PI)",
  "Order Document (Oracle)",
  "Delivery Note (DN)",
  "System Invoice (SI)",
  "Commercial Invoice (CI)",
  "Packing List (PL)",
  "Insurance Certificate",
  "Inspection Certificate",
  "Certificate of Origin (Dubai Chamber)",
  "Shahada Mansha (Ministry of Economy)",
  "MOFA Document",
  "Customs Export Declaration (ED) / Bayan Jumrki",
  "BL / TCN / Manifest / Local",
  "Shared with the Customer",
  "Original Customs Exit / Entry Certificate",
  "Oracle PO (Freight / Inspection / Others)",
  "Vendor Invoice (Freight / Inspection / Others)",
  "Invoice Submission To Finance"
];


// ============================================================================
// EVENT DELEGATION SYSTEM
// ============================================================================

const EVENT_HANDLERS = {
  // Main submit button
  'btnSubmitDcl': handleSubmitDcl,
  
  // Modal actions
  'cancelEnhancedSubmit': cancelEnhancedSubmit,
  'confirmEnhancedSubmit': confirmEnhancedSubmit,
  
  // Document selection
  'selectAllDocuments': selectAllDocuments,
  'deselectAllDocuments': deselectAllDocuments,
  
  // Modals
  'closeDocumentDetails': closeDocumentDetails
};

function setupEventDelegation() {
  // Global click handler for all buttons with data-action
  document.addEventListener('click', function(e) {
    const target = e.target.closest('[data-action]');
    if (!target) return;
    
    const action = target.dataset.action;
    const handler = EVENT_HANDLERS[action];
    
    if (handler) {
      e.preventDefault();
      handler(target);
    }
  });
  
  // Handle document checkboxes
  document.addEventListener('change', function(e) {
    if (e.target.matches('[data-doc-checkbox]')) {
      const docId = e.target.dataset.docCheckbox;
      toggleDocumentSelection(docId);
    }
  });
  
  // Handle info buttons
  document.addEventListener('click', function(e) {
    const infoBtn = e.target.closest('[data-show-details]');
    if (infoBtn) {
      e.preventDefault();
      const docId = infoBtn.dataset.showDetails;
      showDocumentDetails(docId);
    }
  });
  
  // Handle PDF actions
  document.addEventListener('click', function(e) {
    const viewBtn = e.target.closest('[data-view-pdf]');
    if (viewBtn) {
      e.preventDefault();
      viewPdf(viewBtn.dataset.viewPdf);
      return;
    }
    
    const downloadBtn = e.target.closest('[data-download-pdf]');
    if (downloadBtn) {
      e.preventDefault();
      downloadPdf(downloadBtn.dataset.downloadPdf);
      return;
    }
  });
  
  // Handle modal backdrop clicks
  document.addEventListener('click', function(e) {
    if (e.target.matches('[data-modal-close]')) {
      const action = e.target.dataset.modalClose;
      if (action === 'enhancedSubmit') cancelEnhancedSubmit();
      if (action === 'docDetails') closeDocumentDetails();
    }
  });
}



const DOC_TYPE_MAPPING = {
  // Core documents
  "Commercial Invoice": "Commercial Invoice (CI)",
  "Packing List": "Packing List (PL)",
  "Delivery Note": "Delivery Note (DN)",
  "Order Document": "Order Document (Oracle)",
  "System Invoice": "System Invoice (SI)",

  // Certificates
  "Certificate of Origin": "Certificate of Origin (Dubai Chamber)",
  "COO": "Certificate of Origin (Dubai Chamber)",

  // Shipping / Logistics (✅ ALL AWB, BL, Manifest AUTO-MAP)
  "AWB": "BL / TCN / Manifest / Local",
  "BL": "BL / TCN / Manifest / Local",
  "Manifest": "BL / TCN / Manifest / Local",
  "TCN": "BL / TCN / Manifest / Local",

  // Others
  "MOFA": "MOFA Document",
  "Oracle PO": "Oracle PO (Freight / Inspection / Others)",
  "Vendor Invoice": "Vendor Invoice (Freight / Inspection / Others)"
};


// ============================================================================
// INITIALIZATION
// ============================================================================


function getPortalToken() {
  const token = document.querySelector("input[name='__RequestVerificationToken']")?.value;
  if (!token) console.warn("⚠ No __RequestVerificationToken found");
  return token;
}


document.addEventListener('DOMContentLoaded', () => {
  console.log('🚀 DCL Review & Submit page loaded');
  
  // Setup event delegation FIRST
  setupEventDelegation();
  
  initializePage();

  const urlParams = new URLSearchParams(window.location.search);
  const dclId = urlParams.get("id");

  if (dclId) {
    document.querySelectorAll("nav#stepIndicators a").forEach(a => {
      const base = a.getAttribute("href").split("?")[0];
      a.href = `${base}?id=${dclId}`;
    });
  }
});

async function initializePage() {
  try {
    // Clear any existing state
    if (STATE.autoRefreshTimer) {
      clearInterval(STATE.autoRefreshTimer);
      STATE.autoRefreshTimer = null;
    }

    const urlParams = new URLSearchParams(window.location.search);
    STATE.dclMasterId = urlParams.get('id');

    if (!STATE.dclMasterId) {
      showToast('error', 'Error', 'DCL ID not found in URL');
      return;
    }

    console.log(`📋 Loading DCL: ${STATE.dclMasterId}`);
    showLoading('Loading DCL information...');

    // Load data
    await loadDclMasterData();
    await loadCustomerMandatoryDocs();
    await loadDclBrands();
    await loadDclContainers();
    await loadOrderNumbersFromLoadingPlan();
    await loadDclDocuments();

    // Process documents for merge
    processDocumentsForMerge();
    initializeSelection();

    // Check for existing merged PDF
    checkForExistingMergedPdf();

    // Render UI
    renderDclHeader();
    renderDocumentsList();
    renderMandatoryDocsChecklist();
    renderMergedPdfSection();
    renderDclInfoSection();
    updateSelectionUI(); // Set initial submit button state (checks mandatory docs)

    // Initialize autosave AFTER loading data
    autoSaveAdditionalComments();

    // 🔒 LOCK FORM IF SUBMITTED - Call AFTER all UI is rendered
    lockFormIfSubmitted();

    hideLoading();
    console.log('✅ Page initialized successfully');

  } catch (error) {
    console.error('❌ Initialization error:', error);
    hideLoading();
    showToast('error', 'Initialization Failed', error.message);
  }
}

// ============================================================================
// DATA LOADING
// ============================================================================

async function loadDclMasterData() {
  const response = await fetch(`/_api/cr650_dcl_masters(${STATE.dclMasterId})?$expand=cr650_dcl_order_dcl_number_cr650_dcl_master`);
  if (!response.ok) throw new Error(`Failed to load DCL: ${response.status}`);

  const data = await response.json();
  STATE.dclMasterData = data;
  STATE.dclNumber = data.cr650_dclnumber;
  console.log('✅ DCL master loaded:', STATE.dclNumber);
}

async function loadCustomerMandatoryDocs() {
  const customerCode = STATE.dclMasterData?.cr650_customercodes;
  const customerName = STATE.dclMasterData?.cr650_customername;

  if (!customerCode && !customerName) {
    console.log('ℹ️ No customer code or name on DCL — skipping mandatory docs lookup');
    STATE.customerMandatoryDocs = [];
    return;
  }

  try {
    let url;

    // Prefer customer code (unique) over customer name
    if (customerCode) {
      const safeCode = customerCode.replace(/'/g, "''");
      url = `/_api/cr650_updated_dcl_customers?$filter=cr650_customercodes eq '${safeCode}'&$select=cr650_mandatorydocuments&$top=1`;
      console.log(`🔍 Looking up customer by code: ${customerCode}`);
    } else {
      const safeName = customerName.replace(/'/g, "''");
      url = `/_api/cr650_updated_dcl_customers?$filter=cr650_customername eq '${safeName}'&$select=cr650_mandatorydocuments&$top=1`;
      console.log(`🔍 Looking up customer by name: ${customerName}`);
    }

    const response = await fetch(url, {
      headers: { 'Accept': 'application/json', 'OData-Version': '4.0' }
    });

    if (!response.ok) {
      console.error('❌ Failed to load customer mandatory docs:', response.status);
      STATE.customerMandatoryDocs = [];
      return;
    }

    const data = await response.json();
    const customer = (data.value || [])[0];

    if (!customer || !customer.cr650_mandatorydocuments) {
      console.log('ℹ️ No mandatory documents configured for customer:', customerCode || customerName);
      STATE.customerMandatoryDocs = [];
      return;
    }

    // Parse comma-separated string: "PI,CI,PL,Others:custom text"
    const rawDocs = customer.cr650_mandatorydocuments.split(',').map(s => s.trim()).filter(Boolean);

    STATE.customerMandatoryDocs = rawDocs.map(code => {
      if (code.startsWith('Others:') || code === 'Others') {
        const customText = code.startsWith('Others:') ? code.substring(7).trim() : '';
        return { code: code, label: customText || 'Other Document', search: ['other'], isOther: true };
      }
      const config = MANDATORY_DOC_CONFIG[code];
      return {
        code: code,
        label: config ? config.label : code,
        search: config ? config.search : [code.toLowerCase()],
        isOther: false
      };
    });

    console.log(`✅ Loaded ${STATE.customerMandatoryDocs.length} mandatory docs for ${customerCode || customerName}:`, STATE.customerMandatoryDocs);

  } catch (error) {
    console.error('❌ Error loading customer mandatory docs:', error);
    STATE.customerMandatoryDocs = [];
  }
}

/* ============================================================
   LOCK FORM IF SUBMITTED
   ============================================================ */
/* ============================================================
   LOCK FORM IF SUBMITTED
   ============================================================ */
function lockFormIfSubmitted() {
  try {
    const status = (STATE.dclMasterData?.cr650_status || "").toLowerCase();

    if (status !== "submitted" || window.TEST_MODE === true) {
      console.log("📝 Review page is editable - status:", STATE.dclMasterData?.cr650_status || "none");
      return;
    }

    console.log("🔒 Locking review page - DCL status is 'Submitted'");

    // 1. Disable all document selection checkboxes
    document.querySelectorAll('input[type="checkbox"][id^="doc-checkbox-"]').forEach(checkbox => {
      checkbox.disabled = true;
      checkbox.style.cursor = "not-allowed";
      checkbox.style.opacity = "0.5";
    });

    // 2. Hide selection control buttons
    const selectionButtons = document.querySelectorAll('button[data-action="selectAllDocuments"], button[data-action="deselectAllDocuments"]');
    selectionButtons.forEach(btn => {
      btn.style.display = "none";
    });

    // 3. Disable and hide the submit button
    const submitBtn = document.getElementById("btnSubmitDcl");
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.style.display = "none";
    }

    // 4. Hide the entire submit card
    const submitCard = document.querySelector(".submit-final-card");
    if (submitCard) {
      submitCard.style.display = "none";
    }

    // 5. Disable additional comments textarea
    const commentsInput = document.getElementById("additionalCommentsInput");
    if (commentsInput) {
      commentsInput.disabled = true;
      commentsInput.readOnly = true;
      commentsInput.style.cursor = "not-allowed";
      commentsInput.style.opacity = "0.6";
      commentsInput.style.backgroundColor = "#f5f5f5";
      commentsInput.style.pointerEvents = "none";
    }

    // 6. Make document items non-interactive (gray out)
    document.querySelectorAll('.doc-list-item').forEach(item => {
      item.style.opacity = "0.7";
      item.style.cursor = "default";
      item.style.pointerEvents = "none"; // Disable clicking on items
    });

    // 7. KEEP info buttons clickable (re-enable them specifically)
    document.querySelectorAll('button[data-show-details]').forEach(btn => {
      btn.style.pointerEvents = "auto"; // Re-enable pointer events
      btn.style.cursor = "pointer";
      btn.style.opacity = "1";
    });

    // 8. KEEP merged PDF buttons clickable (View & Download)
    const mergedPdfCard = document.querySelector('.merged-pdf-card');
    if (mergedPdfCard) {
      // Re-enable the entire merged PDF section
      mergedPdfCard.style.pointerEvents = "auto";

      // Ensure View and Download buttons are clickable
      const viewBtn = mergedPdfCard.querySelector('button[data-view-pdf]');
      const downloadBtn = mergedPdfCard.querySelector('button[data-download-pdf]');
      
      if (viewBtn) {
        viewBtn.style.pointerEvents = "auto";
        viewBtn.style.cursor = "pointer";
        viewBtn.style.opacity = "1";
        viewBtn.disabled = false;
      }

      if (downloadBtn) {
        downloadBtn.style.pointerEvents = "auto";
        downloadBtn.style.cursor = "pointer";
        downloadBtn.style.opacity = "1";
        downloadBtn.disabled = false;
      }
    }

    // 9. Show locked banner
    showLockedBanner();

    // 10. Update selection counter to show it's locked
    const counter = document.getElementById("selectionCounter");
    if (counter) {
      counter.style.backgroundColor = "#dc2626"; // Red
      counter.innerHTML = "🔒 Locked - Already Submitted";
    }

    console.log("✅ Review page fully locked (Merged PDF and info buttons remain clickable)");

  } catch (lockError) {
    console.error("❌ Error while locking review page:", lockError);
    // Don't throw - just log the error so page still loads
  }
}

function showLockedBanner() {
  try {
    const wizard = document.getElementById("dclWizard");
    const dclHeader = document.querySelector(".dcl-header");

    if (!wizard || !dclHeader) {
      console.warn("Wizard container or DCL header not found");
      return;
    }

    // Check if banner already exists
    if (document.getElementById("lockedBanner")) {
      console.log("Banner already exists");
      return;
    }

    const banner = document.createElement("div");
    banner.id = "lockedBanner";
    banner.style.cssText = `
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%);
            color: white;
            padding: 16px 24px;
            border-radius: 12px;
            margin-bottom: 20px;
            margin-top: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            box-shadow: 0 4px 12px rgba(220, 38, 38, 0.3);
            border: 2px solid #fca5a5;
        `;

    banner.innerHTML = `
            <i class="fas fa-lock" style="font-size: 24px;"></i>
            <div style="flex: 1;">
                <div style="font-size: 16px; font-weight: 700; margin-bottom: 4px;">
                    DCL Locked - Read Only Mode
                </div>
                <div style="font-size: 14px; opacity: 0.95;">
                    This DCL has been submitted. You can view and download the merged PDF, but no changes can be made.
                </div>
            </div>
        `;

    // Insert banner RIGHT AFTER the dcl-header
    dclHeader.insertAdjacentElement('afterend', banner);

    console.log("✅ Locked banner inserted after DCL header");

  } catch (bannerError) {
    console.error("Error creating locked banner:", bannerError);
  }
}

async function loadDclDocuments() {
  const response = await fetch(
    `/_api/cr650_dcl_documents?$filter=_cr650_dcl_number_value eq ${STATE.dclMasterId}&$orderby=createdon desc`
  );
  if (!response.ok) throw new Error(`Failed to load documents: ${response.status}`);

  const data = await response.json();
  STATE.allDocuments = data.value || [];

  console.log(`✅ Loaded ${STATE.allDocuments.length} documents`);
}

// Debounce helper function
function debounce(fn, delay) {
  let timeout;
  return function (...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => fn.apply(this, args), delay);
  };
}


// ============================================================================
// DOCUMENT PROCESSING (CORE BUSINESS LOGIC)
// ============================================================================

function autoSaveAdditionalComments() {
  const input = document.getElementById("additionalCommentsInput");
  if (!input) return;

  // Remove any existing listeners first
  input.removeEventListener("input", input._saveHandler);

  // Create a new handler that captures the current DCL ID
  input._saveHandler = debounce(() => {
    // Use the current STATE.dclMasterId at the time of saving
    const currentDclId = STATE.dclMasterId;
    const comments = input.value.trim();

    if (!currentDclId) {
      console.warn("No DCL ID available for saving");
      return;
    }

    console.log(`Saving comments for DCL: ${currentDclId}`);

    shell.ajaxSafePost({
      type: "PATCH",
      url: `/_api/cr650_dcl_masters(${currentDclId})`,
      contentType: "application/json",
      data: JSON.stringify({
        cr650_additionalcomments: comments || ""
      }),
      headers: {
        "If-Match": "*"
      }
    }).done(function () {
      console.log(`✔ Auto-saved comments for DCL: ${currentDclId}`);
      // Update the local state as well
      if (STATE.dclMasterData) {
        STATE.dclMasterData.cr650_additionalcomments = comments;
      }
    }).fail(function (xhr) {
      console.error(`❌ Auto-save failed for DCL ${currentDclId}:`, xhr.responseText);
    });
  }, 600);

  // Attach the new handler
  input.addEventListener("input", input._saveHandler);
}

// Also make sure to define safeAjax globally if you want consistency
if (typeof window.safeAjax === 'undefined' && typeof shell !== 'undefined') {
  window.safeAjax = shell.ajaxSafePost;
}

function processDocumentsForMerge() {
  console.log('🔄 Processing documents for merge...');

  // Step 1: Exclude merged PDFs (to avoid recursion)
  const nonMergedDocs = STATE.allDocuments.filter(doc => {
    const type = (doc.cr650_doc_type || '').toLowerCase();
    return type !== 'merged pdf' && type !== 'merged dcl package';
  });

  console.log(`📋 After excluding merged PDFs: ${nonMergedDocs.length} documents`);

  // Step 2: Group by document type
  const docsByType = {};
  nonMergedDocs.forEach(doc => {
    const type = doc.cr650_doc_type || 'Other Documents';
    if (!docsByType[type]) docsByType[type] = [];
    docsByType[type].push(doc);
  });

  // Step 3: For each type, select latest per language
  const selectedDocs = [];

  Object.entries(docsByType).forEach(([type, docs]) => {
    // Group by language within this type
    const byLanguage = {};

    docs.forEach(doc => {
      const lang = doc.cr650_documentlanguage || 'en';
      if (!byLanguage[lang]) byLanguage[lang] = [];
      byLanguage[lang].push(doc);
    });

    // For each language, take the latest (already sorted by timestamp desc)
    Object.values(byLanguage).forEach(langDocs => {
      const latest = langDocs[0]; // First item is latest due to $orderby desc
      selectedDocs.push(latest);
    });
  });

  console.log(`📋 After selecting latest per type+language: ${selectedDocs.length} documents`);

  // Step 4: Sort documents according to business rules
  STATE.documentsForMerge = selectedDocs.sort((a, b) => {
    const typeA = a.cr650_doc_type || 'Other Documents';
    const typeB = b.cr650_doc_type || 'Other Documents';

    const priorityA = CONFIG.DOCUMENT_ORDER[typeA] || 999;
    const priorityB = CONFIG.DOCUMENT_ORDER[typeB] || 999;

    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }

    // Same priority - sort by createdon (newest first)
    const timeA = new Date(a.createdon || 0).getTime();
    const timeB = new Date(b.createdon || 0).getTime();
    return timeB - timeA;
  });

  console.log('✅ Documents ordered for merge:', STATE.documentsForMerge.map(d => ({
    type: d.cr650_doc_type,
    language: d.cr650_documentlanguage,
    timestamp: d.cr650_uploadtimestamp
  })));
}


// ============================================================================
// 🆕 DOCUMENT SELECTION MANAGEMENT
// ============================================================================

function initializeSelection() {
  // By default, select all documents
  STATE.selectedDocuments.clear();
  STATE.documentsForMerge.forEach(doc => {
    STATE.selectedDocuments.add(doc.cr650_dcl_documentid);
  });
  console.log(`✅ Initialized with ${STATE.selectedDocuments.size} documents selected`);
}

function toggleDocumentSelection(documentId) {
  if (STATE.selectedDocuments.has(documentId)) {
    STATE.selectedDocuments.delete(documentId);
  } else {
    STATE.selectedDocuments.add(documentId);
  }
  updateSelectionUI();
}

function selectAllDocuments() {
  STATE.selectedDocuments.clear();
  STATE.documentsForMerge.forEach(doc => {
    STATE.selectedDocuments.add(doc.cr650_dcl_documentid);
  });
  updateSelectionUI();
}

function deselectAllDocuments() {
  STATE.selectedDocuments.clear();
  updateSelectionUI();
}

function updateSelectionUI() {
  const total = STATE.documentsForMerge.length;
  const selected = STATE.selectedDocuments.size;

  // Update checkboxes and grayed-out state
  STATE.documentsForMerge.forEach(doc => {
    const checkbox = document.getElementById(`doc-checkbox-${doc.cr650_dcl_documentid}`);
    const docItem = document.getElementById(`doc-item-${doc.cr650_dcl_documentid}`);

    if (checkbox) {
      checkbox.checked = STATE.selectedDocuments.has(doc.cr650_dcl_documentid);
    }

    if (docItem) {
      if (STATE.selectedDocuments.has(doc.cr650_dcl_documentid)) {
        docItem.classList.remove('doc-deselected');
      } else {
        docItem.classList.add('doc-deselected');
      }
    }
  });

  // Update counter
  const counterEl = document.getElementById('selectionCounter');
  if (counterEl) {
    counterEl.textContent = `${selected} of ${total} selected`;
  }

  // Update submit button — also check mandatory docs
  const submitBtn = document.getElementById('btnSubmitDcl');
  if (submitBtn) {
    const mandatoryCheck = validateMandatoryDocuments();
    if (selected === 0) {
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span class="btn-icon"><i class="fas fa-ban"></i></span><span class="btn-text">No Documents Selected</span>';
      submitBtn.classList.add('btn-disabled');
    } else if (!mandatoryCheck.valid) {
      submitBtn.disabled = true;
      submitBtn.innerHTML = `<span class="btn-icon"><i class="fas fa-lock"></i></span><span class="btn-text">${mandatoryCheck.missing.length} Mandatory Doc${mandatoryCheck.missing.length > 1 ? 's' : ''} Missing</span>`;
      submitBtn.classList.add('btn-disabled');
    } else {
      submitBtn.disabled = false;
      submitBtn.innerHTML = `<span class="btn-icon"><i class="fas fa-check-circle"></i></span><span class="btn-text">Submit for Export</span><span class="btn-arrow"><i class="fas fa-arrow-right"></i></span>`;
      submitBtn.classList.remove('btn-disabled');
    }
  }
}

function getSelectedDocuments() {
  return STATE.documentsForMerge.filter(doc =>
    STATE.selectedDocuments.has(doc.cr650_dcl_documentid)
  );
}

function getDeselectedDocuments() {
  return STATE.documentsForMerge.filter(doc =>
    !STATE.selectedDocuments.has(doc.cr650_dcl_documentid)
  );
}

// 🆕 Document Details Modal
function showDocumentDetails(documentId) {
  const doc = STATE.documentsForMerge.find(d => d.cr650_dcl_documentid === documentId);
  if (!doc) return;

  const docType = doc.cr650_doc_type || 'Unknown';
  const language = doc.cr650_documentlanguage || 'en';
  const langLabel = language.toLowerCase() === 'ar' ? 'Arabic' : 'English';
  const timestamp = formatDateTime(doc.cr650_uploadtimestamp || doc.createdon);
  const charge = doc.cr650_chargeamount ? `${doc.cr650_currencycode || 'USD'} ${doc.cr650_chargeamount.toFixed(2)}` : 'N/A';
  const remarks = doc.cr650_remarks || 'None';
  const folder = doc.cr650_customfoldername || 'Default';
  const url = doc.cr650_documenturl || 'N/A';
  const fileExt = doc.cr650_file_extention || 'pdf';

  const modalHTML = `
    <div class="modal-overlay" id="docDetailsModal" style="display: flex !important; ...">
      <div data-modal-close="docDetails" style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; cursor: pointer;"></div>
      <div style="background: white; border-radius: 12px; max-width: 600px; width: 90%; z-index: 2; position: relative; box-shadow: 0 25px 50px rgba(0,0,0,0.5); max-height: 90vh; overflow-y: auto;">
        <div style="display: flex; justify-content: space-between; padding: 1.5rem; border-bottom: 2px solid #e5e7eb;">
          <h3 style="margin: 0; font-size: 1.5rem; font-weight: 600;">
            <i class="fas fa-file-alt"></i> Document Details
          </h3>
          <button data-action="closeDocumentDetails" style="background: none; border: none; font-size: 1.5rem; cursor: pointer;">
            <i class="fas fa-times"></i>
          </button>
        </div>
        <div style="padding: 1.5rem;">
          <div style="display: grid; gap: 1rem;">
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-tag"></i> Type:</strong>
              <span>${escapeHtml(docType)}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-language"></i> Language:</strong>
              <span>${langLabel}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-clock"></i> Uploaded:</strong>
              <span>${timestamp}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-dollar-sign"></i> Charge:</strong>
              <span>${charge}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-file-pdf"></i> File Type:</strong>
              <span>${fileExt.toUpperCase()}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-folder"></i> Folder:</strong>
              <span>${escapeHtml(folder)}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-comment"></i> Remarks:</strong>
              <span style="white-space: pre-wrap;">${escapeHtml(remarks)}</span>
            </div>
            <div style="display: grid; grid-template-columns: 140px 1fr; gap: 1rem; padding: 0.75rem; background: #f3f4f6; border-radius: 8px;">
              <strong style="color: #374151;"><i class="fas fa-link"></i> URL:</strong>
              <a href="${url}" target="_blank" style="color: #3b82f6; word-break: break-all;">${url}</a>
            </div>
          </div>
        </div>
        <div style="display: flex; justify-content: flex-end; padding: 1.5rem; border-top: 2px solid #e5e7eb;">
          <button data-action="closeDocumentDetails" class="btn btn-primary">
            <i class="fas fa-check"></i> Close
          </button>
        </div>
      </div>
    </div>
  `;

  const existing = document.getElementById('docDetailsModal');
  if (existing) existing.remove();

  document.body.insertAdjacentHTML('beforeend', modalHTML);
}

function closeDocumentDetails() {
  const modal = document.getElementById('docDetailsModal');
  if (modal) modal.remove();
}

function formatDateTime(dateString) {
  if (!dateString) return "-";
  try {
    const d = new Date(dateString);
    return d.toLocaleString("en-GB", {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch {
    return "-";
  }
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

function checkForExistingMergedPdf() {
  const mergedDoc = STATE.allDocuments.find(doc => {
    const type = (doc.cr650_doc_type || '').toLowerCase();
    return type === 'merged pdf' || type === 'merged dcl package';
  });

  if (mergedDoc) {
    STATE.mergedPdf = mergedDoc;
    console.log('✅ Found existing merged PDF:', mergedDoc.cr650_documenturl);
  } else {
    STATE.mergedPdf = null;
    console.log('ℹ️ No merged PDF found yet');
  }
}


function getPathForFlow(doc) {
  // ALWAYS return the original portal path stored in Dataverse
  return doc.cr650_documenturl;
}


// ============================================================================
// UI RENDERING
// ============================================================================

function renderDclHeader() {
  const headerEl = document.getElementById('dclNumber');
  if (headerEl && STATE.dclNumber) {
    headerEl.textContent = STATE.dclNumber;
  }
}



function renderDocumentsList() {
  const container = document.getElementById('documentsListContainer');
  if (!container) return;

  if (STATE.documentsForMerge.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="fas fa-inbox"></i>
        <p>No documents to merge yet.</p>
        <p class="hint">Upload documents or generate CI/PL/LP first.</p>
      </div>
    `;
    return;
  }

  const total = STATE.documentsForMerge.length;
  const selected = STATE.selectedDocuments.size;

  let html = `
    <div class="documents-header" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 2px solid #e5e7eb; flex-wrap: wrap; gap: 1rem;">
      <div style="display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;">
        <h3 style="margin: 0; font-size: 1.25rem;">Documents to be Merged</h3>
        <span id="selectionCounter" style="background: #006633; color: white; padding: 0.375rem 0.875rem; border-radius: 999px; font-size: 0.875rem; font-weight: 600;">${selected} of ${total} selected</span>
      </div>
      <div style="display: flex; gap: 0.5rem;">
        <button data-action="selectAllDocuments" style="background: none; border: none; color: #006633; cursor: pointer; padding: 0.5rem 0.75rem; border-radius: 8px; font-weight: 500;">
          <i class="fas fa-check-square"></i> Select All
        </button>
        <button data-action="deselectAllDocuments" style="background: none; border: none; color: #006633; cursor: pointer; padding: 0.5rem 0.75rem; border-radius: 8px; font-weight: 500;">
          <i class="fas fa-square"></i> Deselect All
        </button>
      </div>
    </div>
    <div class="documents-list">
  `;

  STATE.documentsForMerge.forEach((doc, index) => {
    const docType = doc.cr650_doc_type || 'Unknown';
    const language = doc.cr650_documentlanguage || 'en';
    const langLabel = language.toLowerCase() === 'ar' ? '(AR)' : '(EN)';
    const timestamp = formatTimeAgo(doc.createdon);
    const isSelected = STATE.selectedDocuments.has(doc.cr650_dcl_documentid);
    const deselectedClass = isSelected ? '' : 'doc-deselected';

    html += `
      <div class="doc-list-item ${deselectedClass}" id="doc-item-${doc.cr650_dcl_documentid}" style="display: flex; align-items: center; padding: 1rem; background: white; border: 2px solid #e5e7eb; border-radius: 12px; margin-bottom: 0.75rem; gap: 1rem; transition: all 0.3s;">
        <div style="display: flex; align-items: center;">
          <input 
            type="checkbox" 
            id="doc-checkbox-${doc.cr650_dcl_documentid}"
            data-doc-checkbox="${doc.cr650_dcl_documentid}"
            ${isSelected ? 'checked' : ''}
            style="width: 20px; height: 20px; cursor: pointer;"
          />
        </div>
        <div style="min-width: 36px; height: 36px; border-radius: 50%; background: linear-gradient(135deg, #006633, #00884a); color: white; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 0.95rem;">${index + 1}</div>
        <div style="flex: 1; min-width: 0;">
          <div style="font-weight: 600; color: #374151; margin-bottom: 0.25rem; display: flex; align-items: center; gap: 0.5rem;">
            <i class="fas fa-file-pdf" style="color: #ef4444;"></i>
            ${escapeHtml(docType)} ${langLabel}
          </div>
          <div style="font-size: 0.85rem; color: #6b7280;">${timestamp}</div>
        </div>
        <button data-show-details="${doc.cr650_dcl_documentid}" style="background: none; border: 2px solid #e5e7eb; color: #006633; width: 36px; height: 36px; border-radius: 50%; cursor: pointer; display: flex; align-items: center; justify-content: center;">
          <i class="fas fa-info-circle" style="font-size: 1.1rem;"></i>
        </button>
      </div>
    `;
  });

  html += `</div>`;

  container.innerHTML = html;
}


// ============================================================================
// MANDATORY DOCUMENTS CHECKLIST (Customer-Specific)
// ============================================================================

function renderMandatoryDocsChecklist() {
  const container = document.getElementById('mandatoryDocsChecklistContainer');
  if (!container) return;

  // If no mandatory docs configured, hide the section
  if (!STATE.customerMandatoryDocs || STATE.customerMandatoryDocs.length === 0) {
    container.innerHTML = '';
    container.style.display = 'none';
    return;
  }

  container.style.display = 'block';

  const customerName = STATE.dclMasterData?.cr650_customername || 'Customer';

  // Check each mandatory doc against uploaded documents (cr650_dcl_documents)
  let uploadedCount = 0;
  const totalRequired = STATE.customerMandatoryDocs.length;
  const docItems = STATE.customerMandatoryDocs.map(doc => {
    const isUploaded = isMandatoryDocPresent(doc.search);
    if (isUploaded) uploadedCount++;
    return { ...doc, isUploaded };
  });

  const allComplete = uploadedCount === totalRequired;
  const progressPct = totalRequired > 0 ? Math.round((uploadedCount / totalRequired) * 100) : 0;

  let html = `
    <div class="card mandatory-checklist-card" style="background: white; border-radius: 12px; padding: 0; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; border: 2px solid ${allComplete ? '#10b981' : '#f59e0b'};">
      <!-- Header -->
      <div style="padding: 1.25rem 1.5rem; background: ${allComplete ? 'linear-gradient(135deg, #ecfdf5, #d1fae5)' : 'linear-gradient(135deg, #fffbeb, #fef3c7)'}; border-bottom: 2px solid ${allComplete ? '#a7f3d0' : '#fde68a'};">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 0.75rem;">
          <div style="display: flex; align-items: center; gap: 0.75rem;">
            <i class="fas ${allComplete ? 'fa-check-circle' : 'fa-clipboard-list'}" style="font-size: 1.4rem; color: ${allComplete ? '#059669' : '#d97706'};"></i>
            <div>
              <h4 style="margin: 0; font-size: 1.1rem; font-weight: 700; color: #1f2937;">Mandatory Documents</h4>
              <p style="margin: 0; font-size: 0.825rem; color: #6b7280;">Required for <strong>${escapeHtml(customerName)}</strong></p>
            </div>
          </div>
          <span style="background: ${allComplete ? '#059669' : '#d97706'}; color: white; padding: 0.375rem 0.875rem; border-radius: 999px; font-size: 0.825rem; font-weight: 700;">
            ${uploadedCount} / ${totalRequired} Complete
          </span>
        </div>
        <!-- Progress bar -->
        <div style="margin-top: 0.75rem; height: 6px; background: ${allComplete ? '#a7f3d0' : '#fde68a'}; border-radius: 3px; overflow: hidden;">
          <div style="height: 100%; width: ${progressPct}%; background: ${allComplete ? '#059669' : '#d97706'}; border-radius: 3px; transition: width 0.5s ease;"></div>
        </div>
      </div>
      <!-- Checklist items -->
      <div style="padding: 1rem 1.5rem;">
  `;

  docItems.forEach(doc => {
    const statusIcon = doc.isUploaded
      ? '<i class="fas fa-check-circle" style="color: #059669; font-size: 1.15rem;"></i>'
      : '<i class="fas fa-times-circle" style="color: #dc2626; font-size: 1.15rem;"></i>';

    const statusBadge = doc.isUploaded
      ? '<span style="background: #ecfdf5; color: #059669; padding: 0.2rem 0.6rem; border-radius: 4px; font-size: 0.75rem; font-weight: 600;">Uploaded</span>'
      : '<span style="background: #fef2f2; color: #dc2626; padding: 0.2rem 0.6rem; border-radius: 4px; font-size: 0.75rem; font-weight: 600;">Missing</span>';

    const displayName = doc.isOther
      ? `Other: ${escapeHtml(doc.label)}`
      : escapeHtml(doc.label);

    const shortCode = doc.isOther ? '' : `<span style="color: #9ca3af; font-size: 0.8rem; margin-left: 0.5rem;">(${escapeHtml(doc.code)})</span>`;

    html += `
      <div class="mandatory-doc-row" style="display: flex; align-items: center; gap: 0.75rem; padding: 0.625rem 0.5rem; border-bottom: 1px solid #f3f4f6; transition: background 0.2s;">
        ${statusIcon}
        <div style="flex: 1; min-width: 0;">
          <span style="font-size: 0.9rem; font-weight: 500; color: ${doc.isUploaded ? '#374151' : '#991b1b'};">${displayName}</span>${shortCode}
        </div>
        ${statusBadge}
      </div>
    `;
  });

  // Remove last border
  html += `
      </div>
  `;

  // If not all complete, show warning footer
  if (!allComplete) {
    const missingCount = totalRequired - uploadedCount;
    html += `
      <div style="padding: 1rem 1.5rem; background: #fef2f2; border-top: 2px solid #fecaca;">
        <div style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.5rem;">
          <i class="fas fa-exclamation-triangle" style="color: #dc2626;"></i>
          <strong style="color: #991b1b; font-size: 0.9rem;">${missingCount} document${missingCount > 1 ? 's' : ''} still required</strong>
        </div>
        <p style="margin: 0; font-size: 0.825rem; color: #7f1d1d; line-height: 1.5;">
          Upload or generate the missing documents before submitting this DCL.
        </p>
        <div style="display: flex; gap: 0.75rem; flex-wrap: wrap; margin-top: 0.75rem;">
          <a href="/DCL-Document-Generator/?id=${STATE.dclMasterId}" style="display: inline-flex; align-items: center; gap: 0.4rem; background: #006633; color: white; padding: 0.5rem 1rem; border-radius: 8px; font-weight: 600; font-size: 0.825rem; text-decoration: none;">
            <i class="fas fa-file-alt"></i> Document Generator
          </a>
          <a href="/Upload_Center_Accruals/?id=${STATE.dclMasterId}" style="display: inline-flex; align-items: center; gap: 0.4rem; background: white; color: #006633; padding: 0.5rem 1rem; border: 2px solid #006633; border-radius: 8px; font-weight: 600; font-size: 0.825rem; text-decoration: none;">
            <i class="fas fa-upload"></i> Upload Center
          </a>
        </div>
      </div>
    `;
  }

  html += `</div>`;
  container.innerHTML = html;
}

function getLatestMergedPdf(allDocs) {
  const mergedFiles = allDocs.filter(d =>
    d.cr650_doc_type === "Merged PDF" &&
    d.cr650_documenturl
  );

  if (mergedFiles.length === 0) return null;

  // Sort by createdon DESC
  mergedFiles.sort((a, b) =>
    new Date(b.createdon) - new Date(a.createdon)
  );

  return mergedFiles[0];
}

function convertPortalUrlToSharePoint(url) {
  if (!url) return "";

  if (url.includes("sharepoint.com")) {
    // It is already a SharePoint file — return as is
    return url;
  }

  // Portal-relative URL — convert to SharePoint library path
  return "https://petrolubegroup.sharepoint.com/sites/Exportsite" + url;
}


function renderMergedPdfSection() {
  const container = document.getElementById('mergedPdfContainer');
  if (!container) return;

  const merged = getLatestMergedPdf(STATE.allDocuments);
  if (!merged) {
    container.innerHTML = `
      <div class="merged-pdf-placeholder">
        <i class="fas fa-file-pdf"></i>
        <p>Merged PDF will appear here after submission</p>
      </div>
    `;
    return;
  }

  const downloadUrl = convertPortalUrlToSharePoint(merged.cr650_documenturl);

  container.innerHTML = `
    <div class="merged-pdf-ready">
      <div class="pdf-icon">
        <i class="fas fa-file-pdf"></i>
      </div>
      <div class="pdf-info">
        <h4>Merged DCL Package</h4>
        <p>Created ${formatTimeAgo(merged.createdon)}</p>
      </div>
      <div class="pdf-actions">
        <button class="btn btn-secondary" data-view-pdf="${downloadUrl}">
          <i class="fas fa-eye"></i> View
        </button>
        <button class="btn btn-secondary" data-download-pdf="${downloadUrl}">
          <i class="fas fa-download"></i> Download
        </button>
      </div>
    </div>
  `;
}



function renderDclInfoSection() {
  if (!STATE.dclMasterData) return;

  const m = STATE.dclMasterData;

  document.getElementById("info_dclNumber").textContent = m.cr650_dclnumber || "—";
  document.getElementById("info_customer").textContent = m.cr650_customername || "—";
  document.getElementById("info_countryport").textContent =
    `${m.cr650_country || ""} / ${m.cr650_destinationport || ""}`;
  document.getElementById("info_order").textContent =
  STATE.loadingPlanOrders.length
    ? STATE.loadingPlanOrders.join(", ")
    : "—";
  document.getElementById("info_incoterms").textContent = m.cr650_incoterms || "—";
  document.getElementById("info_mode").textContent = m.cr650_transportationmode || "—";

  // Load additional comments
  document.getElementById("additionalCommentsInput").value =
    m.cr650_additionalcomments || "";
}

// ============================================================================
// MANDATORY DOCUMENT VALIDATION
// ============================================================================

/**
 * Validates that all mandatory documents (from customer record) exist in the DCL.
 * Returns { valid: true } or { valid: false, missing: [...names] }
 */
function validateMandatoryDocuments() {
  // If no customer mandatory docs configured, validation passes
  if (!STATE.customerMandatoryDocs || STATE.customerMandatoryDocs.length === 0) {
    return { valid: true, missing: [] };
  }

  const missing = [];

  for (const doc of STATE.customerMandatoryDocs) {
    if (!isMandatoryDocPresent(doc.search)) {
      missing.push(doc.label);
    }
  }

  return missing.length === 0
    ? { valid: true, missing: [] }
    : { valid: false, missing };
}

/**
 * Shows a detailed, user-friendly error banner listing every mandatory
 * document that is missing. Displayed inline above the submit button
 * so users can clearly see what needs to be fixed.
 */
function showMandatoryDocError(missingDocs) {
  // Remove any previous validation banner
  const existing = document.getElementById('mandatoryDocValidation');
  if (existing) existing.remove();

  const listItems = missingDocs.map(doc =>
    `<li style="padding:0.5rem 0;display:flex;align-items:center;gap:0.75rem;">
      <i class="fas fa-times-circle" style="color:#dc3545;font-size:1.1rem;flex-shrink:0;"></i>
      <span>${escapeHtml(doc)}</span>
    </li>`
  ).join('');

  const bannerHTML = `
    <div id="mandatoryDocValidation" style="
      background: #fff5f5;
      border: 2px solid #feb2b2;
      border-left: 5px solid #dc3545;
      border-radius: 12px;
      padding: 1.5rem 2rem;
      margin-bottom: 1.5rem;
      animation: fadeIn 0.3s ease;
    ">
      <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:1rem;">
        <i class="fas fa-exclamation-triangle" style="color:#dc3545;font-size:1.5rem;"></i>
        <h4 style="margin:0;font-size:1.15rem;font-weight:700;color:#991b1b;">
          Cannot Submit — Mandatory Documents Missing
        </h4>
      </div>
      <p style="margin:0 0 0.75rem;color:#7f1d1d;font-size:0.95rem;line-height:1.6;">
        The following <strong>${missingDocs.length}</strong> mandatory document${missingDocs.length > 1 ? 's are' : ' is'} required before this DCL can be submitted for export.
        Please upload or generate ${missingDocs.length > 1 ? 'them' : 'it'} in the <strong>Document Generator</strong> or <strong>Upload Center</strong> steps, then return here to submit.
      </p>
      <ul style="list-style:none;padding:0;margin:0 0 1rem;">
        ${listItems}
      </ul>
      <div style="display:flex;gap:0.75rem;flex-wrap:wrap;">
        <a href="/DCL-Document-Generator/?id=${STATE.dclMasterId}" style="
          display:inline-flex;align-items:center;gap:0.5rem;
          background:#006633;color:white;padding:0.625rem 1.25rem;
          border-radius:8px;font-weight:600;font-size:0.9rem;text-decoration:none;
        ">
          <i class="fas fa-file-alt"></i> Go to Document Generator
        </a>
        <a href="/Upload_Center_Accruals/?id=${STATE.dclMasterId}" style="
          display:inline-flex;align-items:center;gap:0.5rem;
          background:white;color:#006633;padding:0.625rem 1.25rem;
          border:2px solid #006633;border-radius:8px;font-weight:600;font-size:0.9rem;text-decoration:none;
        ">
          <i class="fas fa-upload"></i> Go to Upload Center
        </a>
      </div>
    </div>
  `;

  // Insert above the submit section
  const submitSection = document.querySelector('.submit-section-wrapper');
  if (submitSection) {
    submitSection.insertAdjacentHTML('beforebegin', bannerHTML);
    // Scroll to the banner so the user sees it
    document.getElementById('mandatoryDocValidation')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  } else {
    // Fallback: show as toast
    showToast('error', 'Mandatory Documents Missing',
      `Please upload these documents before submitting: ${missingDocs.join(', ')}`);
  }
}

/**
 * Clears any visible mandatory-document error banner.
 */
function clearMandatoryDocError() {
  const banner = document.getElementById('mandatoryDocValidation');
  if (banner) banner.remove();
}

// ============================================================================
// SUBMIT HANDLER
// ============================================================================

async function handleSubmitDcl() {
  console.log('🔘 Submit button clicked');

  // 🔒 Prevent submission if already submitted
  const status = (STATE.dclMasterData?.cr650_status || "").toLowerCase();
  if (status === "submitted") {
    showToast('error', 'Already Submitted', 'This DCL has already been submitted and cannot be modified.');
    return;
  }

  if (STATE.isSubmitting) {
    console.log('⏳ Submission already in progress');
    return;
  }

  // Clear any previous validation errors
  clearMandatoryDocError();

  const selected = getSelectedDocuments();
  const deselected = getDeselectedDocuments();

  if (selected.length === 0) {
    showToast('warning', 'No Documents Selected', 'Please select at least one document to include in the merged PDF package.');
    return;
  }

  // Validate mandatory documents exist in the DCL
  const validation = validateMandatoryDocuments();
  if (!validation.valid) {
    console.warn('⚠️ Mandatory documents missing:', validation.missing);
    showMandatoryDocError(validation.missing);
    return;
  }

  // Show enhanced confirmation modal
  showEnhancedConfirmation(selected, deselected);
}

function showEnhancedConfirmation(selectedDocs, deselectedDocs) {
  let selectedHTML = selectedDocs.map(doc => {
    const docType = doc.cr650_doc_type || 'Unknown';
    const language = doc.cr650_documentlanguage || 'en';
    const langLabel = language.toLowerCase() === 'ar' ? '(AR)' : '(EN)';
    return `<li style="padding: 0.75rem 1rem; display: flex; align-items: center; gap: 0.75rem; background: rgba(16, 185, 129, 0.05); border-bottom: 1px solid #e5e7eb;"><i class="fas fa-check-circle" style="color: #10b981; font-size: 1.1rem;"></i> ${escapeHtml(docType)} ${langLabel}</li>`;
  }).join('');

  let deselectedHTML = deselectedDocs.length > 0
    ? deselectedDocs.map(doc => {
      const docType = doc.cr650_doc_type || 'Unknown';
      const language = doc.cr650_documentlanguage || 'en';
      const langLabel = language.toLowerCase() === 'ar' ? '(AR)' : '(EN)';
      return `<li style="padding: 0.75rem 1rem; display: flex; align-items: center; gap: 0.75rem; background: rgba(239, 68, 68, 0.05); border-bottom: 1px solid #e5e7eb; color: #6b7280; text-decoration: line-through;"><i class="fas fa-times-circle" style="color: #ef4444; font-size: 1.1rem;"></i> ${escapeHtml(docType)} ${langLabel}</li>`;
    }).join('')
    : '';

  // Check if any deselected docs are mandatory (using customer-specific search terms)
  const deselectedMandatory = deselectedDocs.filter(doc => {
    const docType = (doc.cr650_doc_type || '').toLowerCase();
    return (STATE.customerMandatoryDocs || []).some(md =>
      md.search.some(term => docType.includes(term.toLowerCase()))
    );
  });

  const mandatoryWarningHTML = deselectedMandatory.length > 0
    ? `<div style="background:#fffbeb;border:2px solid #fbbf24;border-radius:8px;padding:1rem;margin-bottom:1.5rem;">
        <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.5rem;">
          <i class="fas fa-exclamation-triangle" style="color:#d97706;font-size:1.1rem;"></i>
          <strong style="color:#92400e;font-size:0.95rem;">Warning: Mandatory documents excluded</strong>
        </div>
        <p style="margin:0;color:#92400e;font-size:0.875rem;line-height:1.5;">
          You have deselected ${deselectedMandatory.length} mandatory document${deselectedMandatory.length > 1 ? 's' : ''}:
          <strong>${deselectedMandatory.map(d => escapeHtml(d.cr650_doc_type)).join(', ')}</strong>.
          These will not appear in the merged PDF package.
        </p>
      </div>`
    : '';

  const modalHTML = `
    <div id="enhancedConfirmModal" style="display: flex !important; position: fixed !important; top: 0 !important; left: 0 !important; right: 0 !important; bottom: 0 !important; width: 100vw !important; height: 100vh !important; background: rgba(0, 0, 0, 0.7) !important; z-index: 999999 !important; align-items: center !important; justify-content: center !important; opacity: 1 !important; visibility: visible !important;">
      <div data-modal-close="enhancedSubmit" style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; z-index: 1; cursor: pointer;"></div>
      <div style="background: white !important; border-radius: 12px; max-width: 600px; width: 90%; max-height: 90vh; z-index: 2 !important; position: relative !important; box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5) !important; overflow-y: auto;">
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 1.5rem; border-bottom: 2px solid #e5e7eb;">
          <h3 style="margin: 0; font-size: 1.5rem; font-weight: 600;">Confirm Document Merge</h3>
          <button data-action="cancelEnhancedSubmit" style="background: none; border: none; font-size: 1.5rem; cursor: pointer;">
            <i class="fas fa-times"></i>
          </button>
        </div>
        <div style="padding: 1.5rem;">
          ${mandatoryWarningHTML}
          <p style="margin-bottom: 1.5rem; color: #6b7280;">You are about to merge the following documents:</p>
          
          <div style="margin-bottom: 1.5rem;">
            <h4 style="font-size: 1rem; font-weight: 600; margin-bottom: 0.75rem; display: flex; align-items: center; gap: 0.5rem;">
              <i class="fas fa-check-circle" style="color: #10b981;"></i>
              Documents to Merge (${selectedDocs.length})
            </h4>
            <ul style="list-style: none; padding: 0; margin: 0; border: 2px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
              ${selectedHTML}
            </ul>
          </div>
          
          ${deselectedDocs.length > 0 ? `
            <div style="margin-bottom: 1.5rem;">
              <h4 style="font-size: 1rem; font-weight: 600; margin-bottom: 0.75rem; display: flex; align-items: center; gap: 0.5rem;">
                <i class="fas fa-times-circle" style="color: #ef4444;"></i>
                Excluded Documents (${deselectedDocs.length})
              </h4>
              <ul style="list-style: none; padding: 0; margin: 0; border: 2px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
                ${deselectedHTML}
              </ul>
            </div>
          ` : ''}
          
          <ul style="list-style: none; padding: 0; margin: 1rem 0 0 0;">
            <li style="padding: 0.5rem 0; color: #6b7280;"><i class="fas fa-clock"></i> Processing time: 30-90 seconds</li>
            <li style="padding: 0.5rem 0; color: #6b7280;"><i class="fas fa-sync"></i> Auto-refresh every 5 seconds</li>
            <li style="padding: 0.5rem 0; color: #6b7280;"><i class="fas fa-file-pdf"></i> PDF will appear when ready</li>
          </ul>
        </div>
        <div style="display: flex; gap: 1rem; justify-content: flex-end; padding: 1.5rem; border-top: 2px solid #e5e7eb;">
          <button data-action="cancelEnhancedSubmit" class="btn btn-secondary">
            <i class="fas fa-times"></i> Cancel
          </button>
          <button data-action="confirmEnhancedSubmit" class="btn btn-primary">
            <i class="fas fa-check"></i> Yes, Submit ${selectedDocs.length} Document${selectedDocs.length !== 1 ? 's' : ''}
          </button>
        </div>
      </div>
    </div>
  `;

  const existing = document.getElementById('enhancedConfirmModal');
  if (existing) existing.remove();

  document.body.insertAdjacentHTML('beforeend', modalHTML);
  document.body.style.overflow = 'hidden';
}

function cancelEnhancedSubmit() {
  const modal = document.getElementById('enhancedConfirmModal');
  if (modal) {
    modal.remove();
    document.body.style.overflow = '';
  }
}

async function confirmEnhancedSubmit() {
  try {
    // Disable submit button immediately
    const submitBtn = document.getElementById('btnSubmitDcl');
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.style.opacity = '0.5';
      submitBtn.style.cursor = 'not-allowed';
    }

    cancelEnhancedSubmit();

    STATE.isSubmitting = true;

    // Show progress overlay
    showSubmissionProgress();

    // Step 1: Validate documents
    updateSubmissionStep('step-validate', 'active');
    const recheck = validateMandatoryDocuments();
    if (!recheck.valid) {
      updateSubmissionStep('step-validate', 'error');
      throw new Error(
        'Mandatory documents missing: ' + recheck.missing.join(', ') +
        '. Please upload them before submitting.'
      );
    }
    await new Promise(resolve => setTimeout(resolve, 400));
    updateSubmissionStep('step-validate', 'complete');

    // Step 2: Merge PDF
    updateSubmissionStep('step-merge', 'active');

    const selectedDocs = STATE.documentsForMerge.filter(d => STATE.selectedDocuments.has(d.cr650_dcl_documentid));

    if (selectedDocs.length === 0) {
      updateSubmissionStep('step-merge', 'error');
      throw new Error('No documents selected for merge');
    }

    console.log(`📋 Submitting ${selectedDocs.length} documents for merge...`);

    const mergeResult = await triggerPdfMerge(selectedDocs);

    if (!mergeResult.success) {
      updateSubmissionStep('step-merge', 'error');
      throw new Error(mergeResult.message || 'PDF merge failed');
    }

    updateSubmissionStep('step-merge', 'complete');

    // Step 3: Update DCL status
    updateSubmissionStep('step-submit', 'active');

    const additionalComments = document.getElementById('additionalCommentsInput')?.value || '';
    await updateDclStatus('Submitted', additionalComments);

    updateSubmissionStep('step-submit', 'complete');

    // Success - wait a moment for user to see completion
    await new Promise(resolve => setTimeout(resolve, 1500));

    // Hide progress modal
    hideSubmissionProgress();

    // Show success toast
    showToast('success', 'Success!', 'DCL submitted successfully for export.');

    // Start polling for the merged PDF
    STATE.previousMergedPdfTime = STATE.mergedPdf?.createdon || null;
    startPollingForMergedPdf();

    // Lock the form
    lockFormIfSubmitted();

    console.log('✅ Enhanced submission complete');

  } catch (error) {
    console.error('❌ Submission error:', error);

    // Hide progress modal
    hideSubmissionProgress();

    // Re-enable submit button
    const submitBtn = document.getElementById('btnSubmitDcl');
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.style.opacity = '1';
      submitBtn.style.cursor = 'pointer';
    }

    // Show descriptive error
    const userMessage = error.message.includes('Mandatory documents missing')
      ? error.message
      : error.message.includes('No documents selected')
        ? 'No documents were selected for merge. Please select the documents you want to include and try again.'
        : error.message.includes('PDF merge failed') || error.message.includes('merge')
          ? 'The PDF merge process failed. This may be a temporary issue — please wait a moment and try again. If the problem persists, contact your administrator.'
          : error.message.includes('Failed to update status')
            ? 'The documents were merged but the DCL status could not be updated. Please refresh the page and check if the submission went through.'
            : 'An unexpected error occurred during submission: ' + error.message + '. Please try again or contact your administrator if the issue persists.';
    showToast('error', 'Submission Failed', userMessage);

  } finally {
    STATE.isSubmitting = false;
  }
}

function toLowerSafe(value) {
  if (value === null || value === undefined) return "";
  return String(value).toLowerCase();
}

// ============================================================================
// HELPER FUNCTIONS FOR SUBMISSION
// ============================================================================

async function triggerPdfMerge(selectedDocuments) {
  console.log('🔄 Triggering PDF merge...');

  try {
    // Call the Power Automate flow
    const result = await callPowerAutomateFlow();

    if (result.success) {
      console.log('✅ PDF merge triggered successfully');
      return {
        success: true,
        message: 'PDF merge initiated'
      };
    } else {
      console.error('❌ PDF merge failed:', result.message);
      return {
        success: false,
        message: result.message || 'PDF merge failed'
      };
    }
  } catch (error) {
    console.error('❌ Error triggering PDF merge:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

async function updateDclStatus(status, comments = '') {
  console.log(`📝 Updating DCL status to: ${status}`);

  try {
    const updateData = {
      cr650_status: status
    };

    // Include comments if provided
    if (comments) {
      updateData.cr650_additionalcomments = comments;
    }

    // Add submission timestamp
    if (status === 'Submitted') {
      updateData.cr650_submitteddate = new Date().toISOString();
    }

    // Use shell.ajaxSafePost for portal authentication
    return new Promise((resolve, reject) => {
      shell.ajaxSafePost({
        type: 'PATCH',
        url: `/_api/cr650_dcl_masters(${STATE.dclMasterId})`,
        contentType: 'application/json',
        data: JSON.stringify(updateData),
        headers: {
          'If-Match': '*'
        }
      })
        .done(function (response) {
          console.log('✅ DCL status updated successfully');

          // Update local state
          if (STATE.dclMasterData) {
            STATE.dclMasterData.cr650_status = status;
            if (comments) {
              STATE.dclMasterData.cr650_additionalcomments = comments;
            }
            if (status === 'Submitted') {
              STATE.dclMasterData.cr650_submitteddate = new Date().toISOString();
            }
          }

          resolve(true);
        })
        .fail(function (xhr, textStatus, errorThrown) {
          console.error('❌ Error updating DCL status:', xhr.responseText);
          reject(new Error(`Failed to update status: ${xhr.status} - ${xhr.responseText}`));
        });
    });

  } catch (error) {
    console.error('❌ Error updating DCL status:', error);
    throw error;
  }
}
// ============================================================================
// POWER AUTOMATE INTEGRATION
// ============================================================================
async function loadDclContainers() {
  const response = await fetch(
    `/_api/cr650_dcl_containers?$filter=_cr650_dcl_number_value eq ${STATE.dclMasterId}`
  );

  if (!response.ok) {
    console.error("Failed to load containers:", response.status, await response.text());
    STATE.containers = [];
    STATE.containerCounts = {
      qty20ft: 0,
      qty40ft: 0,
      tankerCount: 0,
      truckCount: 0
    };
    return;
  }

  const data = await response.json();
  STATE.containers = data.value || [];

  // Initialize counts
  STATE.containerCounts = {
    qty20ft: 0,
    qty40ft: 0,
    tankerCount: 0,
    truckCount: 0
  };

  STATE.containers.forEach(container => {
    // ✅ Get NUMERIC type code (1, 2, 7, 8)
    const type = container.cr650_container_type;

    // Container Type Codes:
    // 1 = 20ft Container
    // 2 = 40ft Container
    // 3 = 40ft High Cube
    // 4 = ISO Tank Container
    // 5 = Flexi Bag 20ft
    // 6 = Flexi Bag 40ft
    // 7 = Bulk Tanker
    // 8 = Truck

    if (type === 1 || type === 5) {
      STATE.containerCounts.qty20ft++;
    }
    if (type === 2 || type === 3 || type === 4 || type === 6) {
      STATE.containerCounts.qty40ft++;
    }
    if (type === 7) {
      STATE.containerCounts.tankerCount++;
    }
    if (type === 8) {
      STATE.containerCounts.truckCount++;
    }
  });
}

async function loadOrderNumbersFromLoadingPlan() {
  if (!STATE.dclMasterId) {
    STATE.loadingPlanOrders = [];
    return;
  }

  try {
    // First, let's see what fields are available
    const testResponse = await fetch(
      `/_api/cr650_dcl_loading_plans?$top=1`
    );
    
    if (testResponse.ok) {
      const testData = await testResponse.json();
      console.log("📋 Sample loading plan record:", testData.value[0]);
      // This will show us the actual field names
    }
  } catch (e) {
    console.log("Test query failed:", e);
  }

  // Now try the actual query with corrected field name
  const response = await fetch(
    `/_api/cr650_dcl_loading_plans` +
    `?$select=cr650_ordernumber` +
    `&$filter=_cr650_dcl_number_value eq ${STATE.dclMasterId}`  // Simplified field name
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error("❌ Failed to load loading plan orders", errorText);
    STATE.loadingPlanOrders = [];
    return;
  }

  const data = await response.json();

  STATE.loadingPlanOrders = [
    ...new Set(
      (data.value || [])
        .map(r => r.cr650_ordernumber)
        .filter(Boolean)
        .map(o => String(o).trim())
    )
  ];

  console.log("✅ Orders from Loading Plan:", STATE.loadingPlanOrders);
}

// ============================================================================
// LOAD DCL BRANDS
// ============================================================================

async function loadDclBrands() {
  if (!STATE.dclMasterId) {
    STATE.dclBrands = [];
    return;
  }

  const response = await fetch(
    `/_api/cr650_dcl_brands` +
    `?$select=cr650_brand` +
    `&$filter=_cr650_dcl_number_value eq ${STATE.dclMasterId}`
  );

  if (!response.ok) {
    console.error("❌ Failed to load brands", await response.text());
    STATE.dclBrands = [];
    return;
  }

  const data = await response.json();

  STATE.dclBrands = (data.value || [])
    .map(b => (b.cr650_brand || "").trim())
    .filter(Boolean);

  console.log("✅ Brands loaded:", STATE.dclBrands);
}



function getBrandString() {
  return STATE.dclBrands.length
    ? STATE.dclBrands.join(", ")
    : "—";
}


// Helper function to format dates
function formatDate(dateString) {
  if (!dateString) return 'N/A';
  try {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch {
    return dateString;
  }
}

function buildChecklistFromDclMaster(master) {
  if (!master) return {};

  // Execution Month-Year
  let executionMonthYear = "";
  const baseDate = master.cr650_loadingdate || master.cr650_submitteddate || master.createdon;
  if (baseDate) {
    executionMonthYear = formatDate(baseDate, "MMM yyyy");
  }

  // Normalize fields for checkbox comparisons
  const payment = (master.cr650_paymentterms || "").toLowerCase();
  const inco = (master.cr650_incoterms || "").toLowerCase();
  const mode = (master.cr650_transportationmode || "").toLowerCase();

  // Combine country and port
  const countryPort = [
    master.cr650_country,
    master.cr650_destinationport
  ].filter(Boolean).join("/");

  return {
    // HEADER FIELDS
    ExecutionMonthYear: executionMonthYear,
    Brand: getBrandString(),

    DCLNo: master.cr650_dclnumber || "",
    CustomerName: master.cr650_customername || "",
    CINumber: master.cr650_ci_number || "",
    CountryPortOfDischarge: countryPort || "",  // FIXED: Combined field
    OrderNumber: STATE.loadingPlanOrders.length
      ? STATE.loadingPlanOrders.join(", ")
      : "N/A",

    // PAYMENT TERMS CHECKBOXES
    LC: payment.includes("lc") ? "☑" : "☐",
    CAD: payment.includes("cad") ? "☑" : "☐",
    Cash: payment.includes("cash") ? "☑" : "☐",
    Credit: payment.includes("credit") ? "☑" : "☐",

    // INCOTERMS CHECKBOXES
    CIF_Checked: inco.includes("cif") ? "☑" : "☐",
    CFR_Checked: inco.includes("cfr") ? "☑" : "☐",
    EXW_Checked: inco.includes("exw") || inco.includes("ex-works") ? "☑" : "☐",
    FOB_Checked: inco.includes("fob") ? "☑" : "☐",

    // TRANSPORTATION MODE CHECKBOXES  
    ByRoad_Checked: mode.includes("road") ? "☑" : "☐",
    ByAir_Checked: mode.includes("air") ? "☑" : "☐",
    ByLaunch_Checked: mode.includes("launch") ? "☑" : "☐",
    BySea_Checked: mode.includes("sea") ? "☑" : "☐",

    // VEHICLE/CONTAINER COUNTS (FROM STATE.containerCounts)
    TankerCount: String(STATE.containerCounts?.tankerCount || 0),
    TruckCount: String(STATE.containerCounts?.truckCount || 0),
    Qty20ft: String(STATE.containerCounts?.qty20ft || 0),
    Qty40ft: String(STATE.containerCounts?.qty40ft || 0),

    // DATE FIELDS (FORMATTED)
    ARD: formatDate(master.cr650_actualreadinessdate),
    ActualLoadingDate: formatDate(master.cr650_loadingdate),
    LCDateOfIssue: formatDate(master.cr650_lcissuedate),
    LCLatestDateOfShipment: formatDate(master.cr650_lclatestshipmentdate),
    BankSubmissionDate: formatDate(master.cr650_banksubmissiondate),
    LCDateOfExpiry: formatDate(master.cr650_lcexpirydate)
  };
}

function normalize(str) {
  return (str || "")
    .toLowerCase()
    .replace(/[\s()\/\-.]/g, "")
    .replace(/,/g, "")
    .trim();
}


function getDocumentStatus(activityName) {
  const normActivity = normalize(activityName);

  const realDocs = STATE.allDocuments.filter(d =>
    (d.cr650_doc_type || "").toLowerCase() !== "merged pdf"
  );

  for (const doc of realDocs) {
    const raw = doc.cr650_doc_type || "";

    // ✅ Direct normalize match
    if (normalize(raw) === normActivity) return "Yes";

    // ✅ Dynamic family detection
    for (const key in DOC_TYPE_MAPPING) {
      if (raw.toLowerCase().includes(key.toLowerCase())) {
        if (normalize(DOC_TYPE_MAPPING[key]) === normActivity) {
          return "Yes";
        }
      }
    }
  }

  return "No";
}


function getDocumentRemarks(activityName) {
  const normActivity = normalize(activityName);
  const collected = [];

  const realDocs = STATE.allDocuments.filter(d =>
    (d.cr650_doc_type || "").toLowerCase() !== "merged pdf"
  );

  for (const doc of realDocs) {
    const raw = doc.cr650_doc_type || "";
    const remark = (doc.cr650_remarks || "").trim();

    if (!remark) continue;

    // ✅ Direct match
    if (normalize(raw) === normActivity) {
      collected.push(remark);
      continue;
    }

    // ✅ Dynamic family match
    for (const key in DOC_TYPE_MAPPING) {
      if (raw.toLowerCase().includes(key.toLowerCase())) {
        if (normalize(DOC_TYPE_MAPPING[key]) === normActivity) {
          collected.push(remark);
        }
      }
    }
  }

  // ✅ REMOVE TRUE DUPLICATES + ✅ NEW LINE FORMAT (NO DASH)
  const unique = [...new Set(collected)];
  return unique.length ? unique.join("\n") : "";
}


async function callPowerAutomateFlow() {
  console.log('📤 Calling Power Automate flow...');

  const selectedDocs = getSelectedDocuments();

  if (selectedDocs.length === 0) {
    throw new Error('No documents selected for merge');
  }

  console.log(`📋 Sending ${selectedDocs.length} selected documents to flow`);

  // -----------------------------------------------------
  // 1) Create checklist key/value dictionary for Word
  // -----------------------------------------------------
  const checklistMapped = {};

  STATE.checklistMaster.forEach(name => {
    const clean = name
      .replace(/[\s()\/\-.]/g, "")
      .replace(/,/g, "");

    const status = getDocumentStatus(name);
    const remarks = getDocumentRemarks(name);

    checklistMapped[clean + "_Status"] = status;
    checklistMapped[clean + "_Remarks"] = remarks;
  });

  // Build checklist data
  const checklist = buildChecklistFromDclMaster(STATE.dclMasterData);

  // Build template master
  const templateMaster = {
    ExportExecutiveName:
      STATE.dclMasterData?.cr650_submitter_name ||
      STATE.dclMasterData?.cr650_submitter_email ||
      "Unknown User",
    DCLOpenDate:
      STATE.dclMasterData?.["createdon@OData.Community.Display.V1.FormattedValue"] || "",
    DCLClosureDate:
      STATE.dclMasterData?.["cr650_submitteddate@OData.Community.Display.V1.FormattedValue"] ||
      "",
    CINumber: STATE.dclMasterData?.cr650_ci_number || "",
    // ✅ OrderNumber removed - it's already in checklist!
    CountryPortOfDischarge: STATE.dclMasterData?.cr650_destinationport || "",
    AdditionalComments: STATE.dclMasterData?.cr650_additionalcomments || ""
  };

  // ✅ NEW: Merge all three into templateData
  const templateData = {
    ...checklist,
    ...checklistMapped,
    ...templateMaster
  };

  // ✅ NEW: Calculate completeness
  const filledCount = Object.values(templateData).filter(v =>
    v !== null && v !== undefined && v !== ""
  ).length;
  const totalCount = Object.keys(templateData).length;
  const percentage = Math.round((filledCount / totalCount) * 100);

  // Build documents array
  const documentsArray = selectedDocs.map((doc, index) => ({
    documentId: doc.cr650_dcl_documentid,
    documentUrl: getPathForFlow(doc),
    documentType: doc.cr650_doc_type,
    language: doc.cr650_documentlanguage || null,
    orderIndex: index + 1
  }));

  // -----------------------------------------------------
  // 2) Build payload - ✅ NEW FORMAT
  // -----------------------------------------------------
  const payload = {
    // ✅ NEW: metadata object
    metadata: {
      extractedAt: new Date().toISOString(),
      dclId: STATE.dclMasterId,
      dclNumber: STATE.dclNumber,
      completeness: {
        total: totalCount,
        filled: filledCount,
        percentage: percentage,
        missing: totalCount - filledCount
      }
    },

    // ✅ NEW: templateData (all 67 fields in one object)
    templateData: templateData,

    // ✅ NEW: raw data objects
    rawMasterData: STATE.dclMasterData || {},
    rawDocuments: STATE.allDocuments || [],
    rawContainers: [],

    // Keep these for compatibility/additional info
    dclNumber: STATE.dclNumber,
    dclMasterId: STATE.dclMasterId,
    documents: documentsArray,
    totalDocuments: selectedDocs.length,
    requestTimestamp: new Date().toISOString()
  };

  console.log("📋 FINAL Payload:", payload);

  // -----------------------------------------------------
  // 3) Call the flow
  // -----------------------------------------------------
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), CONFIG.FLOW_TIMEOUT);

  try {
    const response = await fetch(CONFIG.POWER_AUTOMATE_FLOW_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      throw new Error(`Flow returned ${response.status}: ${await response.text()}`);
    }

    const result = await response.json();
    console.log('✅ Flow result:', result);

    return {
      success: result.success !== false,
      message: result.message,
      ...result
    };

  } catch (err) {
    clearTimeout(timeoutId);

    if (err.name === 'AbortError') {
      throw new Error("Flow timeout — may still finish in background.");
    }
    throw new Error(`Flow failed: ${err.message}`);
  }
}

// ============================================================================
// POLLING FOR MERGED PDF
// ============================================================================

function startPollingForMergedPdf() {
  console.log('🔄 Starting to poll for merged PDF...');

  let pollCount = 0;
  const maxPolls = 30; // 30 polls * 5 seconds = 2.5 minutes max

  STATE.autoRefreshTimer = setInterval(async () => {
    pollCount++;

    if (pollCount > maxPolls) {
      stopPolling();
      showToast('info', 'Still Processing', 'Merge is taking longer than expected. Please refresh the page manually.');
      return;
    }

    try {
      console.log(`🔄 Poll attempt ${pollCount}/${maxPolls}...`);

      await loadDclDocuments();
      checkForExistingMergedPdf();

      if (STATE.mergedPdf) {

        const newTime = new Date(STATE.mergedPdf.createdon).getTime();
        const oldTime = STATE.previousMergedPdfTime
          ? new Date(STATE.previousMergedPdfTime).getTime()
          : 0;

        if (newTime > oldTime) {
          console.log('✅ NEW merged PDF detected!');
          stopPolling();
          renderMergedPdfSection();
          showToast('success', 'Merge Complete', 'Your merged PDF is ready!');
        } else {
          console.log("⏳ Old merged PDF detected — waiting for new one...");
        }
      }


    } catch (error) {
      console.error('Polling error:', error);
    }
  }, CONFIG.REFRESH_INTERVAL);
}

function stopPolling() {
  if (STATE.autoRefreshTimer) {
    clearInterval(STATE.autoRefreshTimer);
    STATE.autoRefreshTimer = null;
  }
}

// ============================================================================
// PDF ACTIONS
// ============================================================================

function viewPdf(url) {
  window.open(url, "_blank");
}

function downloadPdf(url) {
  const a = document.createElement("a");
  a.href = url + "?download=1";
  a.target = "_blank";
  a.rel = "noopener";
  a.download = "";   // tells browser to download instead of open
  document.body.appendChild(a);
  a.click();
  a.remove();
}



// ============================================================================
// UI UTILITIES
// ============================================================================

function showLoading(message = 'Loading...') {
  const overlay = document.getElementById('loadingOverlay');
  if (!overlay) return;

  const textEl = overlay.querySelector('.loading-text');
  const subtextEl = overlay.querySelector('.loading-subtext');

  if (message.includes('\n')) {
    const [mainText, subText] = message.split('\n');
    if (textEl) textEl.textContent = mainText;
    if (subtextEl) subtextEl.textContent = subText;
  } else {
    if (textEl) textEl.textContent = message;
    if (subtextEl) subtextEl.textContent = 'Please wait...';
  }

  overlay.classList.remove('hidden');
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function showToast(type, title, message) {
  const container = document.getElementById('toastContainer');
  if (!container) {
    alert(`${title}\n\n${message}`);
    return;
  }

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  const icon = type === 'success' ? 'fa-check-circle' :
    type === 'error' ? 'fa-exclamation-circle' :
      type === 'warning' ? 'fa-exclamation-triangle' :
        'fa-info-circle';

  toast.innerHTML = `
    <div class="toast-icon">
      <i class="fas ${icon}"></i>
    </div>
    <div class="toast-content">
      <h4>${title}</h4>
      <p>${message}</p>
    </div>
  `;

  container.appendChild(toast);
  setTimeout(() => toast.classList.add('show'), 10);

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 5000);
}

// NEW: Progress toast for step-by-step feedback
function showProgressToast(message, stepId) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  // Remove old progress toasts
  container.querySelectorAll('.toast-progress').forEach(t => t.remove());

  const toast = document.createElement('div');
  toast.className = 'toast toast-progress toast-info';
  toast.id = `toast-${stepId}`;

  toast.innerHTML = `
    <div class="toast-icon">
      <div class="spinner-small"></div>
    </div>
    <div class="toast-content">
      <h4>Processing</h4>
      <p>${message}</p>
    </div>
  `;

  container.appendChild(toast);
  setTimeout(() => toast.classList.add('show'), 10);
}

// ============================================================================
// SUBMISSION PROGRESS MODAL
// ============================================================================

function showSubmissionProgress() {
  const overlay = document.createElement('div');
  overlay.id = 'submissionProgressOverlay';
  overlay.className = 'submission-progress-overlay';

  overlay.innerHTML = `
    <div class="progress-modal">
      <div class="progress-header">
        <i class="fas fa-file-export"></i>
        <h3>Submitting DCL for Export</h3>
      </div>
      
      <div class="progress-steps">
        <div class="progress-step" id="step-validate">
          <div class="step-icon"><i class="fas fa-circle"></i></div>
          <div class="step-info">
            <h4>Validating Documents</h4>
            <p>Checking all requirements...</p>
          </div>
        </div>
        
        <div class="progress-step" id="step-merge">
          <div class="step-icon"><i class="fas fa-circle"></i></div>
          <div class="step-info">
            <h4>Merging PDF</h4>
            <p>Combining selected documents...</p>
          </div>
        </div>
        
        <div class="progress-step" id="step-submit">
          <div class="step-icon"><i class="fas fa-circle"></i></div>
          <div class="step-info">
            <h4>Updating Status</h4>
            <p>Finalizing submission...</p>
          </div>
        </div>
      </div>
      
      <div class="progress-bar-container">
        <div class="progress-bar-fill" id="submissionProgressBar"></div>
      </div>
      
      <p class="progress-note">
        <i class="fas fa-info-circle"></i>
        This process takes 30-90 seconds. Please do not close this window.
      </p>
    </div>
  `;

  document.body.appendChild(overlay);
  setTimeout(() => overlay.classList.add('show'), 10);
}

function updateSubmissionStep(stepId, status = 'active') {
  const step = document.getElementById(stepId);
  if (!step) return;

  // Remove previous states
  step.classList.remove('active', 'complete', 'error');
  step.classList.add(status);

  const icon = step.querySelector('.step-icon i');
  if (status === 'complete') {
    icon.className = 'fas fa-check-circle';
  } else if (status === 'error') {
    icon.className = 'fas fa-times-circle';
  } else if (status === 'active') {
    icon.className = 'fas fa-spinner fa-spin';
  }

  // Update progress bar
  const steps = ['step-validate', 'step-merge', 'step-submit'];
  const currentIndex = steps.indexOf(stepId);
  const progress = ((currentIndex + 1) / steps.length) * 100;
  document.getElementById('submissionProgressBar').style.width = progress + '%';
}

function hideSubmissionProgress() {
  const overlay = document.getElementById('submissionProgressOverlay');
  if (overlay) {
    overlay.classList.remove('show');
    setTimeout(() => overlay.remove(), 300);
  }
}

function formatTimeAgo(dateString) {
  if (!dateString) return 'recently';

  try {
    const date = new Date(dateString);
    const seconds = Math.floor((new Date() - date) / 1000);

    if (seconds < 60) return 'just now';
    if (seconds < 3600) return Math.floor(seconds / 60) + ' minutes ago';
    if (seconds < 86400) return Math.floor(seconds / 3600) + ' hours ago';
    if (seconds < 604800) return Math.floor(seconds / 86400) + ' days ago';

    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  } catch {
    return 'recently';
  }
}



// ============================================================================
// MANUAL REFRESH
// ============================================================================

async function refreshPage() {
  try {
    showLoading('Refreshing...');
    await loadDclDocuments();
    processDocumentsForMerge();
    checkForExistingMergedPdf();
    renderDocumentsList();
    renderMandatoryDocsChecklist();
    renderMergedPdfSection();
    updateSelectionUI();
  } catch (error) {
    showToast('error', 'Refresh Failed', error.message);
  } finally {
    hideLoading();
  }
}

// ============================================================================
// EXPORT GLOBALS
// ============================================================================
window.handleSubmitDcl = handleSubmitDcl;
window.confirmEnhancedSubmit = confirmEnhancedSubmit; // 🆕
window.cancelEnhancedSubmit = cancelEnhancedSubmit; // 🆕
window.toggleDocumentSelection = toggleDocumentSelection; // 🆕
window.selectAllDocuments = selectAllDocuments; // 🆕
window.deselectAllDocuments = deselectAllDocuments; // 🆕
window.showDocumentDetails = showDocumentDetails; // 🆕
window.closeDocumentDetails = closeDocumentDetails; // 🆕
window.viewPdf = viewPdf;
window.showSubmissionProgress = showSubmissionProgress;
window.updateSubmissionStep = updateSubmissionStep;
window.hideSubmissionProgress = hideSubmissionProgress;
window.showProgressToast = showProgressToast;
window.downloadPdf = downloadPdf;
window.refreshPage = refreshPage;
window.DCL_STATE = STATE;

console.log('📦 DCL Review & Submit (CORRECTED) loaded');
// ============================================================================