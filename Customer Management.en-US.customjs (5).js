/**
 * ═══════════════════════════════════════════════════════════════════════
 * CUSTOMER MANAGEMENT SYSTEM - POWER PAGES OPTIMIZED
 * JavaScript Module - Customer Models Feature
 * ═══════════════════════════════════════════════════════════════════════
 *
 * CHANGES IN THIS VERSION:
 * ✅ Replaced fixed address types (Consignee, Ship-To, Bill-To) with dynamic Customer Models
 * ✅ Customer Models CRUD operations (Create, Read, Update, Delete)
 * ✅ Customer Models linked to customers via cr650_is_related_to_a_customer
 * ✅ Fetch and display customer models when editing a customer
 * ✅ Updated Excel export to remove old address columns
 * ✅ Previous features maintained: Sales Rep, Country, Payment Terms, COB with "Other" options
 *
 * Version: 5.0.0 - Customer Models Feature
 * Last Updated: February 5, 2026
 * ═══════════════════════════════════════════════════════════════════════
 */

(function () {
    'use strict';


    function safeAjax(options) {
        return new Promise((resolve, reject) => {
            if (!(window.shell && shell.getTokenDeferred)) {
                reject("Anti-forgery token not available");
                return;
            }

            shell.getTokenDeferred().done(function (token) {
                $.ajax({
                    ...options,
                    headers: {
                        ...(options.headers || {}),
                        "__RequestVerificationToken": token
                    },
                    success: resolve,
                    error: reject
                });
            });
        });
    }
    window.safeAjax = safeAjax;

    // ═══════════════════════════════════════════════════════════════════════
    // CONFIGURATION
    // ═══════════════════════════════════════════════════════════════════════
    const CONFIG = {
        entityName: 'cr650_updated_dcl_customers',
        entitySetName: 'cr650_updated_dcl_customers',
        apiPath: '/_api',
        pageSize: 10,
        debugMode: true
    };

    // Customer Models API Configuration
    const CUSTOMER_MODEL_API = '/_api/cr650_dcl_customer_models';
    const CUSTOMER_MODEL_FIELDS = [
        'cr650_dcl_customer_modelid',
        'cr650_model_name',
        'cr650_info',
        '_cr650_dcl_id_value',
        '_cr650_cutomer_id_value',
        'cr650_is_related_to_a_customer',
        'createdon'
    ];

    // ═══════════════════════════════════════════════════════════════════════
    // APPLICATION STATE
    // ═══════════════════════════════════════════════════════════════════════
    const state = {
        customers: [],
        filteredCustomers: [],
        currentPage: 1,
        totalPages: 1,
        isLoading: false,
        editingCustomerId: null,
        currentStep: 1,
        totalSteps: 3,
        customerModels: [],
        pendingCustomerModels: [], // Models to be created when saving a new customer
        editingModelId: null,
        editingPendingModelIndex: null // Index for editing pending models
    };

    // ═══════════════════════════════════════════════════════════════════════
    // COUNTRY SUGGESTIONS (Smart Hints)
    // ═══════════════════════════════════════════════════════════════════════
    const countrySuggestions = {
        'Yemen': {
            ports: ['Aden', 'Hodeidah'],
            paymentTerms: '100% Cash in advance'
        },
        'Lebanon': {
            ports: ['Beirut'],
            paymentTerms: 'Net 150 Days from Shipment Date'
        },
        'Iraq': {
            ports: ['Umm Qasr', 'Basra'],
            paymentTerms: 'Net due in 120 Days'
        },
        'Jordan': {
            ports: ['Aqaba', 'Amman'],
            paymentTerms: 'Net 60 Days from Shipment Date'
        },
        'Turkey': {
            ports: ['Mersin Port', 'Istanbul'],
            paymentTerms: 'LC'
        },
        'United Arab Emirates': {
            ports: ['Jebel Ali', 'Dubai'],
            paymentTerms: 'Net due in 90 Days'
        },
        'Syrian Arab Republic': {
            ports: ['Latakia', 'Tartus'],
            paymentTerms: 'Immediate'
        }
    };


    // ═══════════════════════════════════════════════════════════════════════
    // COUNTRY → ISO-2 STANDARD MAP (CANONICAL)
    // ═══════════════════════════════════════════════════════════════════════
    const countryToISO2 = {
        'Yemen': 'YE',
        'Lebanon': 'LB',
        'Iraq': 'IQ',
        'Jordan': 'JO',
        'Turkey': 'TR',
        'United Arab Emirates': 'AE',
        'Syrian Arab Republic': 'SY',
        'Kuwait': 'KW',
        'Oman': 'OM',
        'Pakistan': 'PK',
        'Somalia': 'SO',
        'South Africa': 'ZA',
        'Sri Lanka': 'LK',
        'Afghanistan': 'AF',
        'Ivory Coast': 'CI',
        'United States': 'US',
        'Bolivia': 'BO'
    };


    // ═══════════════════════════════════════════════════════════════════════
    // HELPER FUNCTIONS
    // ═══════════════════════════════════════════════════════════════════════
    function getElement(selector) {
        return document.querySelector(selector);
    }

    function log(message, data = '') {
        if (CONFIG.debugMode) {
            console.log(`[Customer Management] ${message}`, data);
        }
    }

    function showLoader(show) {
        const loader = getElement('#topLoader');
        if (loader) {
            if (show) {
                loader.classList.remove('hidden');
                state.isLoading = true;
            } else {
                loader.classList.add('hidden');
                state.isLoading = false;
            }
        }
    }

    function showToast(message, type = 'success') {
        const toast = getElement('#toast');
        const icon = getElement('#toastIcon');
        const messageSpan = getElement('#toastMessage');

        if (toast && icon && messageSpan) {
            toast.className = `toast-cm ${type}`;
            icon.className = type === 'success' ? 'fas fa-check-circle' : 'fas fa-exclamation-circle';
            messageSpan.textContent = message;

            toast.classList.remove('hidden');

            setTimeout(() => {
                toast.classList.add('hidden');
            }, 4000);
        }
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function validateCountryCodeSoft() {
        const input = getElement('#country1');
        if (!input || !input.dataset.suggested) return;

        if (
            input.value.trim().toUpperCase() !==
            input.dataset.suggested.toUpperCase()
        ) {
            showToast(
                `⚠️ Suggested country code is "${input.dataset.suggested}", but you used "${input.value}".`,
                'warning'
            );
        }
    }


    function getCountryFlag(country) {
        const flags = {
            'Yemen': '🇾🇪',
            'Lebanon': '🇱🇧',
            'Iraq': '🇮🇶',
            'Jordan': '🇯🇴',
            'Turkey': '🇹🇷',
            'United Arab Emirates': '🇦🇪',
            'Syrian Arab Republic': '🇸🇾'
        };
        return flags[country] || '🌍';
    }

    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    // ═══════════════════════════════════════════════════════════════════════
    // NEW: "OTHER" OPTION HANDLING
    // ═══════════════════════════════════════════════════════════════════════

    /**
     * Handle Country dropdown change - show/hide custom input for "Other"
     */
    function handleCountryChange(e) {
        const country = e.target.value;
        const suggestion = countrySuggestions[country];
        const customInput = getElement('#countryCustom');
        const countryCodeInput = getElement('#country1');

        // Show/hide custom input for "Other" option
        if (customInput) {
            if (country === 'Other') {
                customInput.style.display = 'block';
                customInput.required = true;
            } else {
                customInput.style.display = 'none';
                customInput.required = false;
                customInput.value = '';
            }
        }

        if (countryToISO2[country]) {
            countryCodeInput.value = countryToISO2[country];

            // 🔓 SMART MODE: suggest, don’t force
            countryCodeInput.disabled = false;
            countryCodeInput.dataset.suggested = countryToISO2[country];
        } else {
            countryCodeInput.value = '';
            countryCodeInput.disabled = false;
            delete countryCodeInput.dataset.suggested;
        }


        // Update hints
        const portHint = getElement('#portSuggestion');
        const paymentHint = getElement('#paymentSuggestion');

        if (suggestion && country !== 'Other') {
            if (portHint) portHint.textContent = `💡 Popular ports: ${suggestion.ports.join(', ')}`;
            if (paymentHint) paymentHint.textContent = `💡 Common: ${suggestion.paymentTerms}`;
        } else {
            if (portHint) portHint.textContent = '';
            if (paymentHint) paymentHint.textContent = '';
        }
    }

    /**
     * Handle Payment Terms dropdown change - show/hide custom input for "Other"
     */
    function handlePaymentTermsChange(e) {
        const value = e.target.value;
        const customInput = getElement('#paymentTermsCustom');

        if (customInput) {
            if (value === 'Other') {
                customInput.style.display = 'block';
                customInput.required = true;
            } else {
                customInput.style.display = 'none';
                customInput.required = false;
                customInput.value = '';
            }
        }
    }

    /**
     * Handle COB dropdown change - show/hide custom input for "Other"
     */
    function handleCOBChange(e) {
        const value = e.target.value;
        const customInput = getElement('#cobCustom');

        if (customInput) {
            if (value === 'Other') {
                customInput.style.display = 'block';
                customInput.required = true;
            } else {
                customInput.style.display = 'none';
                customInput.required = false;
                customInput.value = '';
            }
        }
    }

    /**
     * Handle Sales Rep dropdown change - show/hide custom input for "Other"
     */
    function handleSalesRepChange(e) {
        const value = e.target.value;
        const customInput = getElement('#salesRepCustom');

        if (customInput) {
            if (value === 'Other') {
                customInput.style.display = 'block';
                customInput.required = true;
            } else {
                customInput.style.display = 'none';
                customInput.required = false;
                customInput.value = '';
            }
        }
    }

    /**
     * Handle LP Comments checkbox change - show/hide comments textarea
     */
    function handleLpCommentsToggle(e) {
        const isChecked = e.target.checked;
        const commentsContainer = getElement('#lpCommentsContainer');
        const commentsTextarea = getElement('#lpComments');

        if (commentsContainer) {
            commentsContainer.style.display = isChecked ? 'block' : 'none';
        }

        // Clear comments if unchecked
        if (!isChecked && commentsTextarea) {
            commentsTextarea.value = '';
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CUSTOMER MODELS CRUD OPERATIONS
    // ═══════════════════════════════════════════════════════════════════════

    /**
     * Build query URL for fetching customer models by customer ID
     */
    function buildCustomerModelQueryUrl(customerId) {
        const select = '$select=' + encodeURIComponent(CUSTOMER_MODEL_FIELDS.join(','));
        const filter = '$filter=' + encodeURIComponent(`_cr650_cutomer_id_value eq ${customerId}`);
        const orderby = '$orderby=' + encodeURIComponent('createdon asc');
        return `${CUSTOMER_MODEL_API}?${select}&${filter}&${orderby}`;
    }

    /**
     * Normalize customer model data from API response
     */
    function normalizeCustomerModel(r) {
        return {
            id: r.cr650_dcl_customer_modelid || '',
            modelName: r.cr650_model_name || '',
            info: r.cr650_info || '',
            dclId: r._cr650_dcl_id_value || '',
            customerId: r._cr650_cutomer_id_value || '',
            isRelatedToCustomer: r.cr650_is_related_to_a_customer || false,
            createdOn: r.createdon || ''
        };
    }

    /**
     * Fetch customer models for a specific customer
     */
    async function fetchCustomerModels(customerId) {
        try {
            const url = buildCustomerModelQueryUrl(customerId);
            const response = await fetch(url, {
                headers: {
                    'Accept': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });

            if (!response.ok) {
                throw new Error(`API error: ${response.status}`);
            }

            const data = await response.json();
            state.customerModels = (data.value || []).map(normalizeCustomerModel);
            log('Fetched customer models:', state.customerModels);
            return state.customerModels;
        } catch (error) {
            console.error('Error fetching customer models:', error);
            state.customerModels = [];
            return [];
        }
    }

    /**
     * Create a new customer model
     */
    async function createCustomerModel(customerId, modelName, info) {
        const bodyObj = {
            cr650_model_name: modelName,
            cr650_info: info || '',
            cr650_is_related_to_a_customer: true,
            'cr650_cutomer_id@odata.bind': `/cr650_updated_dcl_customers(${customerId})`
        };

        await safeAjax({
            type: 'POST',
            url: CUSTOMER_MODEL_API,
            contentType: 'application/json',
            data: JSON.stringify(bodyObj)
        });

        return { ok: true };
    }

    /**
     * Update an existing customer model
     */
    async function updateCustomerModel(id, modelName, info) {
        const url = `${CUSTOMER_MODEL_API}(${encodeURIComponent(id)})`;
        const bodyObj = {
            cr650_model_name: modelName,
            cr650_info: info || ''
        };

        await safeAjax({
            type: 'PATCH',
            url: url,
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'If-Match': '*'
            },
            data: JSON.stringify(bodyObj)
        });

        return { ok: true };
    }

    /**
     * Delete a customer model
     */
    async function deleteCustomerModel(id) {
        const url = `${CUSTOMER_MODEL_API}(${encodeURIComponent(id)})`;
        await safeAjax({
            type: 'DELETE',
            url: url,
            headers: { 'If-Match': '*' }
        });
        return { ok: true };
    }

    /**
     * Render customer models list in the UI
     */
    function renderCustomerModelsList() {
        const listContainer = getElement('#customerModelsList');
        const infoContainer = getElement('#customerModelsInfo');
        const addBtnContainer = getElement('#addModelBtnContainer');

        if (!listContainer) return;

        // Determine which models to display
        const isNewCustomer = !state.editingCustomerId;
        const modelsToShow = isNewCustomer ? state.pendingCustomerModels : state.customerModels;

        // Always show the add button and list
        if (infoContainer) infoContainer.style.display = 'none';
        if (addBtnContainer) addBtnContainer.style.display = 'block';
        listContainer.style.display = 'block';

        if (modelsToShow.length === 0) {
            const message = isNewCustomer
                ? 'Add customer models here. They will be saved when you save the customer.'
                : 'No customer models defined yet.';
            listContainer.innerHTML = `
                <div class="address-card-cm" style="margin-bottom: var(--space-lg); background: #f9f9f9; border: 1px dashed #ccc; text-align: center; padding: 2rem;">
                    <i class="fas fa-inbox" style="font-size: 2rem; color: var(--text-tertiary); margin-bottom: 0.5rem;"></i>
                    <p style="color: var(--text-secondary); font-size: 0.9375rem;">${message}</p>
                    <p style="color: var(--text-tertiary); font-size: 0.875rem;">Click "Add Customer Model" to create one.</p>
                </div>
            `;
        } else {
            listContainer.innerHTML = modelsToShow.map((model, index) => {
                // For existing customers, use model.id; for new customers, use index
                const modelIdentifier = isNewCustomer ? `pending-${index}` : model.id;
                return `
                <div class="address-card-cm customer-model-item" data-model-id="${modelIdentifier}" style="margin-bottom: var(--space-md); border-left: 4px solid var(--primary);">
                    <div class="address-card-header-cm" style="display: flex; align-items: center; margin-bottom: var(--space-sm); padding-bottom: var(--space-sm); border-bottom: 1px solid #eee;">
                        <i class="fas fa-cube" style="font-size: 1rem; color: var(--primary); margin-right: 0.5rem;"></i>
                        <strong style="flex-grow: 1; font-size: 1rem; color: var(--text-primary); font-weight: 600;">${escapeHtml(model.modelName)}</strong>
                        ${isNewCustomer ? '<span style="background: #fff3cd; color: #856404; padding: 0.125rem 0.5rem; border-radius: 4px; font-size: 0.7rem; margin-right: 0.5rem;">Pending</span>' : ''}
                        <button type="button" class="btn-edit-model-cm" data-model-id="${modelIdentifier}" style="padding: 0.25rem 0.5rem; background: #e8f5ed; color: var(--primary); border: 1px solid var(--primary); border-radius: 4px; font-size: 0.75rem; font-weight: 600; cursor: pointer; margin-right: 0.5rem;">
                            <i class="fas fa-edit"></i> Edit
                        </button>
                        <button type="button" class="btn-delete-model-cm" data-model-id="${modelIdentifier}" style="padding: 0.25rem 0.5rem; background: #fee; color: #c33; border: 1px solid #fcc; border-radius: 4px; font-size: 0.75rem; font-weight: 600; cursor: pointer;">
                            <i class="fas fa-trash"></i> Delete
                        </button>
                    </div>
                    <div style="padding: 0.5rem 0; white-space: pre-wrap; font-size: 0.875rem; color: var(--text-secondary); line-height: 1.6;">
                        ${model.info ? escapeHtml(model.info) : '<em style="color: var(--text-tertiary);">No information provided</em>'}
                    </div>
                </div>
            `}).join('');
        }
    }

    /**
     * Show the add/edit customer model form
     */
    function showCustomerModelForm(modelId = null) {
        const formContainer = getElement('#customerModelForm');
        const formTitle = getElement('#modelFormTitle');
        const modelNameInput = getElement('#modelName');
        const modelInfoInput = getElement('#modelInfo');
        const editingModelIdInput = getElement('#editingModelId');
        const addBtnContainer = getElement('#addModelBtnContainer');

        if (!formContainer) return;

        state.editingModelId = modelId;
        state.editingPendingModelIndex = null;

        if (modelId) {
            // Check if editing a pending model (for new customers)
            if (modelId.toString().startsWith('pending-')) {
                const index = parseInt(modelId.replace('pending-', ''), 10);
                const model = state.pendingCustomerModels[index];
                if (model) {
                    state.editingPendingModelIndex = index;
                    if (formTitle) formTitle.textContent = 'Edit Model';
                    if (modelNameInput) modelNameInput.value = model.modelName;
                    if (modelInfoInput) modelInfoInput.value = model.info;
                    if (editingModelIdInput) editingModelIdInput.value = modelId;
                }
            } else {
                // Edit existing model (for existing customers)
                const model = state.customerModels.find(m => m.id === modelId);
                if (model) {
                    if (formTitle) formTitle.textContent = 'Edit Model';
                    if (modelNameInput) modelNameInput.value = model.modelName;
                    if (modelInfoInput) modelInfoInput.value = model.info;
                    if (editingModelIdInput) editingModelIdInput.value = modelId;
                }
            }
        } else {
            // Add mode
            if (formTitle) formTitle.textContent = 'Add New Model';
            if (modelNameInput) modelNameInput.value = '';
            if (modelInfoInput) modelInfoInput.value = '';
            if (editingModelIdInput) editingModelIdInput.value = '';
        }

        formContainer.style.display = 'block';
        if (addBtnContainer) addBtnContainer.style.display = 'none';
        if (modelNameInput) modelNameInput.focus();
    }

    /**
     * Hide the customer model form
     */
    function hideCustomerModelForm() {
        const formContainer = getElement('#customerModelForm');
        const addBtnContainer = getElement('#addModelBtnContainer');
        const modelNameInput = getElement('#modelName');
        const modelInfoInput = getElement('#modelInfo');
        const editingModelIdInput = getElement('#editingModelId');

        if (formContainer) formContainer.style.display = 'none';
        if (addBtnContainer) addBtnContainer.style.display = 'block';
        if (modelNameInput) modelNameInput.value = '';
        if (modelInfoInput) modelInfoInput.value = '';
        if (editingModelIdInput) editingModelIdInput.value = '';
        state.editingModelId = null;
    }

    /**
     * Save customer model (create or update)
     */
    async function saveCustomerModel() {
        const modelNameInput = getElement('#modelName');
        const modelInfoInput = getElement('#modelInfo');
        const editingModelIdInput = getElement('#editingModelId');

        const modelName = modelNameInput?.value.trim();
        const modelInfo = modelInfoInput?.value.trim();
        const editingId = editingModelIdInput?.value;

        if (!modelName) {
            showToast('Please enter a model name', 'error');
            if (modelNameInput) modelNameInput.focus();
            return;
        }

        // For new customers (no editingCustomerId), store models locally
        if (!state.editingCustomerId) {
            if (state.editingPendingModelIndex !== null) {
                // Update existing pending model
                state.pendingCustomerModels[state.editingPendingModelIndex] = {
                    modelName: modelName,
                    info: modelInfo
                };
                showToast('Model updated', 'success');
            } else {
                // Add new pending model
                state.pendingCustomerModels.push({
                    modelName: modelName,
                    info: modelInfo
                });
                showToast('Model added (will be saved with customer)', 'success');
            }
            renderCustomerModelsList();
            hideCustomerModelForm();
            return;
        }

        // For existing customers, save to API
        showLoader(true);
        try {
            if (editingId && !editingId.startsWith('pending-')) {
                // Update existing model
                await updateCustomerModel(editingId, modelName, modelInfo);
                showToast('Customer model updated successfully', 'success');
            } else {
                // Create new model
                await createCustomerModel(state.editingCustomerId, modelName, modelInfo);
                showToast('Customer model created successfully', 'success');
            }

            // Refresh the models list
            await fetchCustomerModels(state.editingCustomerId);
            renderCustomerModelsList();
            hideCustomerModelForm();
        } catch (error) {
            console.error('Error saving customer model:', error);
            showToast(error?.responseText || 'Failed to save customer model', 'error');
        } finally {
            showLoader(false);
        }
    }

    /**
     * Handle deleting a customer model
     */
    async function handleDeleteCustomerModel(modelId) {
        // Handle pending models for new customers
        if (modelId.toString().startsWith('pending-')) {
            const index = parseInt(modelId.replace('pending-', ''), 10);
            const model = state.pendingCustomerModels[index];
            const modelName = model ? model.modelName : 'this model';

            if (!confirm(`Are you sure you want to delete "${modelName}"?`)) {
                return;
            }

            state.pendingCustomerModels.splice(index, 1);
            showToast('Model removed', 'success');
            renderCustomerModelsList();
            return;
        }

        // Handle existing models for existing customers
        const model = state.customerModels.find(m => m.id === modelId);
        const modelName = model ? model.modelName : 'this model';

        if (!confirm(`Are you sure you want to delete "${modelName}"?\n\nThis action cannot be undone.`)) {
            return;
        }

        showLoader(true);
        try {
            await deleteCustomerModel(modelId);
            showToast('Customer model deleted successfully', 'success');

            // Refresh the models list
            await fetchCustomerModels(state.editingCustomerId);
            renderCustomerModelsList();
        } catch (error) {
            console.error('Error deleting customer model:', error);
            showToast(error?.responseText || 'Failed to delete customer model', 'error');
        } finally {
            showLoader(false);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // POWER PAGES: FORCE HIDE MODAL ON LOAD
    // ═══════════════════════════════════════════════════════════════════════
    function forceHideModal() {
        const modal = getElement('#customerModal');
        if (modal) {
            modal.classList.remove('active');
            modal.style.display = 'none';
            log('Modal force hidden on page load');
        }

        document.documentElement.style.overflow = 'auto';
        document.body.style.overflow = 'auto';

        // Reset Notify Party 2 visibility on close
        const notifyParty2Container = getElement('#notifyParty2Container');
        const addNotifyPartyBtn = getElement('#addNotifyPartyBtn');
        if (notifyParty2Container) notifyParty2Container.style.display = 'none';
        if (addNotifyPartyBtn) addNotifyPartyBtn.style.display = 'block';

        state.editingCustomerId = null;
        state.currentStep = 1;
        state.pendingCustomerModels = [];
        state.editingPendingModelIndex = null;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // INITIALIZATION
    // ═══════════════════════════════════════════════════════════════════════
    function initialize() {
        log('Application initializing...');

        try {
            forceHideModal(); // CRITICAL: Hide modal immediately
            initializeEventListeners();
            loadCustomers();
            log('Application initialized successfully');
        } catch (error) {
            console.error('Initialization error:', error);
            showToast('Failed to initialize application', 'error');
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // EVENT LISTENERS - POWER PAGES OPTIMIZED
    // ═══════════════════════════════════════════════════════════════════════
    function initializeEventListeners() {
        // Search and filter
        const searchInput = getElement('#searchInput');
        const countryFilter = getElement('#countryFilter');

        if (searchInput) {
            searchInput.addEventListener('input', debounce(handleSearch, 300));
        }
        if (countryFilter) {
            countryFilter.addEventListener('change', handleFilter);
        }

        // Action buttons
        const btnAddCustomer = getElement('#btnAddCustomer');
        const btnExportExcel = getElement('#btnExportExcel');

        if (btnAddCustomer) {
            btnAddCustomer.addEventListener('click', openAddCustomerModal);
        }
        if (btnExportExcel) {
            btnExportExcel.addEventListener('click', exportToExcel);
        }

        // Modal controls
        const btnCloseModal = getElement('#btnCloseModal');
        const btnCancel = getElement('#btnCancel');
        const modalOverlay = getElement('.modal-overlay-cm');

        if (btnCloseModal) {
            btnCloseModal.addEventListener('click', closeModal);
        }
        if (btnCancel) {
            btnCancel.addEventListener('click', closeModal);
        }
        if (modalOverlay) {
            modalOverlay.addEventListener('click', closeModal);
        }

        // Wizard navigation
        const btnNextStep = getElement('#btnNextStep');
        const btnPrevStep = getElement('#btnPrevStep');
        const btnSaveCustomer = getElement('#btnSaveCustomer');

        if (btnNextStep) {
            btnNextStep.addEventListener('click', nextStep);
        }
        if (btnPrevStep) {
            btnPrevStep.addEventListener('click', prevStep);
        }
        if (btnSaveCustomer) {
            btnSaveCustomer.addEventListener('click', saveCustomer);
        }

        // Pagination
        const btnPrevPage = getElement('#btnPrevPage');
        const btnNextPage = getElement('#btnNextPage');

        if (btnPrevPage) {
            btnPrevPage.addEventListener('click', () => changePage(-1));
        }
        if (btnNextPage) {
            btnNextPage.addEventListener('click', () => changePage(1));
        }

        // Form interactions
        const btnCheckCode = getElement('#btnCheckCode');
        const country = getElement('#country');
        const paymentTerms = getElement('#paymentTerms');
        const cob = getElement('#cob');


        if (btnCheckCode) {
            btnCheckCode.addEventListener('click', checkCustomerCode);
        }

        // NEW: Event listeners for "Other" option handling
        if (country) {
            country.addEventListener('change', handleCountryChange);
        }
        if (paymentTerms) {
            paymentTerms.addEventListener('change', handlePaymentTermsChange);
        }
        if (cob) {
            cob.addEventListener('change', handleCOBChange);
        }

        const salesRep = getElement('#salesRep');
        if (salesRep) {
            salesRep.addEventListener('change', handleSalesRepChange);
        }

        // LP Comments checkbox toggle
        const hasLpComments = getElement('#hasLpComments');
        if (hasLpComments) {
            hasLpComments.addEventListener('change', handleLpCommentsToggle);
        }

        // POWER PAGES: Event delegation for edit buttons
        document.addEventListener('click', function (e) {
            const editBtn = e.target.closest('.btn-edit-cm');
            if (editBtn) {
                e.preventDefault();
                const customerId = editBtn.getAttribute('data-customer-id');
                if (customerId) {
                    log('Edit button clicked for customer:', customerId);
                    handleEditCustomer(customerId);
                }
            }
        });

        // Notify Party 2 Controls
        const btnAddNotifyParty2 = getElement('#btnAddNotifyParty2');
        const btnRemoveNotifyParty2 = getElement('#btnRemoveNotifyParty2');
        
        if (btnAddNotifyParty2) {
            btnAddNotifyParty2.addEventListener('click', showNotifyParty2);
        }
        if (btnRemoveNotifyParty2) {
            btnRemoveNotifyParty2.addEventListener('click', hideNotifyParty2);
        }

        // POWER PAGES: Event delegation for delete buttons
        document.addEventListener('click', function (e) {
            const deleteBtn = e.target.closest('.btn-delete-cm');
            if (deleteBtn) {
                e.preventDefault();
                const customerId = deleteBtn.getAttribute('data-customer-id');
                if (customerId) {
                    log('Delete button clicked for customer:', customerId);
                    handleDeleteCustomer(customerId);
                }
            }
        });

        // Customer Model Controls
        const btnAddCustomerModel = getElement('#btnAddCustomerModel');
        const btnCancelModelForm = getElement('#btnCancelModelForm');
        const btnSaveModel = getElement('#btnSaveModel');

        if (btnAddCustomerModel) {
            btnAddCustomerModel.addEventListener('click', () => showCustomerModelForm(null));
        }
        if (btnCancelModelForm) {
            btnCancelModelForm.addEventListener('click', hideCustomerModelForm);
        }
        if (btnSaveModel) {
            btnSaveModel.addEventListener('click', saveCustomerModel);
        }

        // POWER PAGES: Event delegation for edit model buttons
        document.addEventListener('click', function (e) {
            const editModelBtn = e.target.closest('.btn-edit-model-cm');
            if (editModelBtn) {
                e.preventDefault();
                const modelId = editModelBtn.getAttribute('data-model-id');
                if (modelId) {
                    log('Edit model button clicked for model:', modelId);
                    showCustomerModelForm(modelId);
                }
            }
        });

        // POWER PAGES: Event delegation for delete model buttons
        document.addEventListener('click', function (e) {
            const deleteModelBtn = e.target.closest('.btn-delete-model-cm');
            if (deleteModelBtn) {
                e.preventDefault();
                const modelId = deleteModelBtn.getAttribute('data-model-id');
                if (modelId) {
                    log('Delete model button clicked for model:', modelId);
                    handleDeleteCustomerModel(modelId);
                }
            }
        });

        log('Event listeners initialized (Power Pages mode)');
    }

    // ═══════════════════════════════════════════════════════════════════════
    // WIZARD STEP MANAGEMENT
    // ═══════════════════════════════════════════════════════════════════════
    function nextStep() {
        if (!validateCurrentStep()) {
            return;
        }

        if (state.currentStep < state.totalSteps) {
            state.currentStep++;
            updateWizardUI();
        }
    }

    function prevStep() {
        if (state.currentStep > 1) {
            state.currentStep--;
            updateWizardUI();
        }
    }

    function updateWizardUI() {
        // Progress indicators (top steps)
        document.querySelectorAll('.step-cm').forEach((step, index) => {
            const stepNumber = index + 1;
            step.classList.remove('active-cm', 'completed-cm');

            if (stepNumber < state.currentStep) {
                step.classList.add('completed-cm');
            } else if (stepNumber === state.currentStep) {
                step.classList.add('active-cm');
            }
        });

        // 🔥 CRITICAL FIX: Force-hide all steps (inline)
        document.querySelectorAll('.form-step-cm').forEach(step => {
            step.classList.remove('active-cm');
            step.style.display = 'none'; // ← override Power Pages inline styles
        });

        // 🔥 Force-show active step
        const activeStep = document.querySelector(
            `.form-step-cm[data-step="${state.currentStep}"]`
        );
        if (activeStep) {
            activeStep.classList.add('active-cm');
            activeStep.style.display = 'block'; // ← THIS FIXES IT
        }

        // Buttons
        const btnPrev = getElement('#btnPrevStep');
        const btnNext = getElement('#btnNextStep');
        const btnSave = getElement('#btnSaveCustomer');

        if (btnPrev) {
            btnPrev.style.display =
                state.currentStep === 1 ? 'none' : 'inline-flex';
        }

        if (state.currentStep === state.totalSteps) {
            if (btnNext) btnNext.style.display = 'none';
            if (btnSave) btnSave.style.display = 'inline-flex';
        } else {
            if (btnNext) btnNext.style.display = 'inline-flex';
            if (btnSave) btnSave.style.display = 'none';
        }

        // 🔽 UX FIX: reset scroll so user actually sees the new step
        const modalBody = document.querySelector('.modal-body-cm');
        if (modalBody) {
            modalBody.scrollTop = 0;
        }
    }


    /**
     * ENHANCED: Validate current step with better error messages
     */
    function validateCurrentStep() {
        const currentStepEl = document.querySelector(`.form-step-cm[data-step="${state.currentStep}"]`);
        if (!currentStepEl) return true;

        const requiredInputs = currentStepEl.querySelectorAll('[required]');
        let isValid = true;
        let firstInvalidField = null;

        requiredInputs.forEach(input => {
            // Skip hidden custom inputs if their parent dropdown isn't set to "Other"
            if (input.id === 'countryCustom' && getElement('#country').value !== 'Other') {
                return;
            }
            if (input.id === 'paymentTermsCustom' && getElement('#paymentTerms').value !== 'Other') {
                return;
            }
            if (input.id === 'cobCustom' && getElement('#cob').value !== 'Other') {
                return;
            }
            if (input.id === 'salesRepCustom' && getElement('#salesRep').value !== 'Other') {
                return;
            }

            if (!input.value || !input.value.trim()) {
                if (!firstInvalidField) {
                    firstInvalidField = input;
                }
                isValid = false;
            }
        });

        if (!isValid && firstInvalidField) {
            firstInvalidField.focus();

            // Get field label for better error message
            const label = currentStepEl.querySelector(`label[for="${firstInvalidField.id}"]`) ||
                firstInvalidField.closest('.form-group-cm')?.querySelector('label');
            const fieldName = label ? label.textContent.replace('*', '').trim() : 'this field';

            showToast(`Please fill in "${fieldName}" before proceeding`, 'error');
        }

        return isValid;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // DATA OPERATIONS - API Calls
    // ═══════════════════════════════════════════════════════════════════════
    async function loadCustomers() {
        try {
            showLoader(true);

            const url = `${CONFIG.apiPath}/${CONFIG.entitySetName}?$select=cr650_updated_dcl_customerid,cr650_customercodes,cr650_customername,cr650_country,cr650_paymentterms,cr650_salesrepresentativename,cr650_destinationport,cr650_notifyparty1,cr650_notifyparty2,cr650_organizationid,cr650_cob,cr650_country_1,cr650_loadingplancomments,cr650_currency,modifiedon&$orderby=createdon desc`;

            const response = await fetch(url, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });

            if (!response.ok) {
                throw new Error(`API error: ${response.status}`);
            }

            const data = await response.json();

            state.customers = data.value || [];
            state.filteredCustomers = [...state.customers];

            updateStatistics();
            renderCustomerTable();

            showToast(`Loaded ${state.customers.length} customers successfully`, 'success');
        } catch (error) {
            console.error('Error loading customers:', error);
            showToast('Failed to load customers', 'error');
            state.customers = [];
            state.filteredCustomers = [];
            renderCustomerTable();
        } finally {
            showLoader(false);
        }
    }

    /**
     * ENHANCED: Save customer with "Other" option handling
     */
    async function saveCustomer() {
        try {
            const form = getElement('#customerForm');
            if (!form || !form.checkValidity()) {
                if (form) form.reportValidity();
                return;
            }

            showLoader(true);

            // Get country value (use custom if "Other" selected)
            let countryValue = getElement('#country').value;
            if (countryValue === 'Other') {
                const customCountry = getElement('#countryCustom');
                countryValue = customCountry ? customCountry.value.trim() : '';
            }

            // Get payment terms value (use custom if "Other" selected)
            let paymentTermsValue = getElement('#paymentTerms').value;
            if (paymentTermsValue === 'Other') {
                const customPaymentTerms = getElement('#paymentTermsCustom');
                paymentTermsValue = customPaymentTerms ? customPaymentTerms.value.trim() : '';
            }

            // Get COB value (use custom if "Other" selected)
            let cobValue = getElement('#cob').value;
            if (cobValue === 'Other') {
                const customCOB = getElement('#cobCustom');
                cobValue = customCOB ? customCOB.value.trim() : '';
            }

            // Get Sales Rep value (use custom if "Other" selected)
            let salesRepValue = getElement('#salesRep').value;
            if (salesRepValue === 'Other') {
                const customSalesRep = getElement('#salesRepCustom');
                salesRepValue = customSalesRep ? customSalesRep.value.trim() : '';
            }

            // Get LP Comments value (only if checkbox is checked)
            const hasLpCommentsChecked = getElement('#hasLpComments')?.checked;
            const lpCommentsValue = hasLpCommentsChecked
                ? (getElement('#lpComments')?.value.trim() || null)
                : null;

            // Customer data (addresses are now managed via Customer Models)
            const customerData = {
                'cr650_customercodes': getElement('#customerCode').value.trim().toUpperCase(),
                'cr650_customername': getElement('#customerName').value.trim(),
                'cr650_country': countryValue,
                'cr650_salesrepresentativename': salesRepValue,
                'cr650_destinationport': getElement('#destinationPort').value.trim(),
                'cr650_currency': getElement('#currency')?.value || null,
                'cr650_paymentterms': paymentTermsValue,
                'cr650_cob': cobValue,
                'cr650_organizationid': getElement('#organizationId').value.trim(),
                'cr650_country_1': getElement('#country1').value.trim(),
                'cr650_notifyparty1': getElement('#notifyParty1')?.value.trim() || null,
                'cr650_notifyparty2': getElement('#notifyParty2')?.value.trim() || null,
                'cr650_loadingplancomments': lpCommentsValue
            };

            let url = `${CONFIG.apiPath}/${CONFIG.entitySetName}`;
            let method = 'POST';
            let isNewCustomer = !state.editingCustomerId;
            const customerCode = getElement('#customerCode').value.trim().toUpperCase();

            if (state.editingCustomerId) {
                url += `(${state.editingCustomerId})`;
                method = 'PATCH';
            }

            // For POST (create): remove null/undefined/empty - Power Pages doesn't accept explicit nulls in POST
            // For PATCH (update): keep empty strings so fields can be cleared, only remove null/undefined
            const cleanedData = Object.fromEntries(
                Object.entries(customerData).filter(([_, v]) => {
                    if (v == null) return false;
                    if (!state.editingCustomerId && v === '') return false;
                    return true;
                })
            );

            log('Saving customer data:', cleanedData);

            const ajaxOptions = {
                type: method,
                url: url,
                contentType: 'application/json',
                data: JSON.stringify(cleanedData),
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            };

            await safeAjax(ajaxOptions);

            validateCountryCodeSoft();

            // If new customer with pending models, find the customer ID and create models
            if (isNewCustomer && state.pendingCustomerModels.length > 0) {
                // Query for the newly created customer by customer code
                const escapedCode = customerCode.replace(/'/g, "''");
                const findUrl = `${CONFIG.apiPath}/${CONFIG.entitySetName}?$filter=cr650_customercodes eq '${escapedCode}'&$select=cr650_updated_dcl_customerid&$top=1`;
                const findResponse = await fetch(findUrl, {
                    headers: {
                        'Accept': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0'
                    }
                });

                if (findResponse.ok) {
                    const findData = await findResponse.json();
                    if (findData.value && findData.value.length > 0) {
                        const newCustomerId = findData.value[0].cr650_updated_dcl_customerid;

                        // Create all pending customer models
                        log('Creating pending customer models:', state.pendingCustomerModels.length);
                        for (const pendingModel of state.pendingCustomerModels) {
                            try {
                                await createCustomerModel(newCustomerId, pendingModel.modelName, pendingModel.info);
                                log('Created model:', pendingModel.modelName);
                            } catch (modelError) {
                                console.error('Error creating model:', pendingModel.modelName, modelError);
                            }
                        }
                    }
                }

                // Clear pending models
                state.pendingCustomerModels = [];
                showToast('Customer and models saved successfully!', 'success');
            } else {
                showToast(
                    state.editingCustomerId
                        ? 'Customer updated successfully!'
                        : 'Customer created successfully!',
                    'success'
                );
            }

            closeModal();
            await loadCustomers();

        } catch (error) {
            console.error('Error saving customer:', error);
            showToast(
                error?.responseText || error?.statusText || 'Failed to save customer',
                'error'
            );
        }
        finally {
            showLoader(false);
        }

    }

    async function checkDuplicateCode(code) {
        try {
            const escapedCode = code.replace(/'/g, "''");
            const url = `${CONFIG.apiPath}/${CONFIG.entitySetName}?$filter=cr650_customercodes eq '${escapedCode}'&$select=cr650_updated_dcl_customerid`;

            const response = await fetch(url, {
                headers: {
                    'Accept': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });

            if (!response.ok) return null;

            const data = await response.json();
            if (data.value && data.value.length > 0) {
                return data.value[0].cr650_updated_dcl_customerid;
            }
            return null;
        } catch (error) {
            console.error('Error checking duplicate:', error);
            return null;
        }
    }

    async function getCustomerById(id) {
        try {
            const data = await safeAjax({
                type: 'GET',
                url: `${CONFIG.apiPath}/${CONFIG.entitySetName}(${id})`,
                headers: {
                    'Accept': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });
            return data;
        } catch (error) {
            console.error('Error fetching customer:', error);
            return null;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CUSTOMER CODE VALIDATION
    // ═══════════════════════════════════════════════════════════════════════
    async function checkCustomerCode() {
        const codeInput = getElement('#customerCode');
        const validationDiv = getElement('#codeValidation');

        if (!codeInput || !validationDiv) return;

        const code = codeInput.value.trim().toUpperCase();

        const codePattern = /^TL\d{5}$/;
        if (!codePattern.test(code)) {
            validationDiv.className = 'validation-message-cm error';
            validationDiv.innerHTML = '<i class="fas fa-times-circle"></i> Invalid format. Must be TL followed by 5 digits (e.g., TL20142)';
            return;
        }

        showLoader(true);
        try {
            const duplicateId = await checkDuplicateCode(code);

            // Duplicate exists and it's not the customer currently being edited
            if (duplicateId && duplicateId !== state.editingCustomerId) {
                validationDiv.className = 'validation-message-cm error';
                validationDiv.innerHTML = '<i class="fas fa-exclamation-triangle"></i> Customer code already exists!';
            } else {
                validationDiv.className = 'validation-message-cm success';
                validationDiv.style.display = 'flex';

                validationDiv.innerHTML = '<i class="fas fa-check-circle"></i> Customer code is valid and available';
            }
        } catch (error) {
            validationDiv.className = 'validation-message-cm error';
            validationDiv.style.display = 'flex';
            validationDiv.innerHTML = '<i class="fas fa-times-circle"></i> Error checking code availability';
        } finally {
            showLoader(false);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // UI RENDERING - POWER PAGES OPTIMIZED
    // ═══════════════════════════════════════════════════════════════════════
    function renderCustomerTable() {
        const tbody = getElement('#customerTableBody');
        if (!tbody) return;

        if (state.filteredCustomers.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="text-align: center; padding: 3rem; color: var(--text-tertiary);">
                        <i class="fas fa-inbox" style="font-size: 3rem; margin-bottom: 1rem; display: block; opacity: 0.5;"></i>
                        <h3 style="margin-bottom: 0.5rem;">No customers found</h3>
                        <p>Try adjusting your search or filters</p>
                    </td>
                </tr>
            `;
            updatePagination();
            return;
        }

        state.totalPages = Math.ceil(state.filteredCustomers.length / CONFIG.pageSize);
        const startIndex = (state.currentPage - 1) * CONFIG.pageSize;
        const endIndex = startIndex + CONFIG.pageSize;
        const pageCustomers = state.filteredCustomers.slice(startIndex, endIndex);

        // POWER PAGES: Using data attributes instead of inline onclick
        tbody.innerHTML = pageCustomers.map(customer => `
            <tr>
                <td><strong class="customer-code-cm">${escapeHtml(customer.cr650_customercodes || 'N/A')}</strong></td>
                <td>${escapeHtml(customer.cr650_customername || 'N/A')}</td>
                <td><span class="country-badge-cm">${getCountryFlag(customer.cr650_country)} ${escapeHtml(customer.cr650_country || 'N/A')}</span></td>
                <td><span class="payment-terms-cm">${escapeHtml(customer.cr650_paymentterms || 'N/A')}</span></td>
                <td>${escapeHtml(customer.cr650_salesrepresentativename || 'N/A')}</td>
                <td class="text-center">
                    <button class="btn-edit-cm" data-customer-id="${customer.cr650_updated_dcl_customerid}" type="button">
                        <i class="fas fa-edit"></i>
                        Edit
                    </button>
                    <button class="btn-delete-cm" data-customer-id="${customer.cr650_updated_dcl_customerid}" type="button">
                        <i class="fas fa-trash-alt"></i>
                        Delete
                    </button>
                </td>
            </tr>
        `).join('');

        updatePagination();
    }

    function updatePagination() {
        const prevBtn = getElement('#btnPrevPage');
        const nextBtn = getElement('#btnNextPage');
        const pageInfo = getElement('#pageInfo');

        if (prevBtn) prevBtn.disabled = state.currentPage === 1;
        if (nextBtn) nextBtn.disabled = state.currentPage >= state.totalPages;
        if (pageInfo) pageInfo.textContent = `Page ${state.currentPage} of ${state.totalPages || 1}`;
    }

    function updateStatistics() {
        const totalEl = getElement('#totalCustomers');
        if (totalEl) totalEl.textContent = state.customers.length;

        const uniqueCountries = new Set(state.customers.map(c => c.cr650_country).filter(Boolean));
        const countriesEl = getElement('#totalCountries');
        if (countriesEl) countriesEl.textContent = uniqueCountries.size;

        const today = new Date().toDateString();
        const updatedToday = state.customers.filter(c => {
            if (!c.modifiedon) return false;
            return new Date(c.modifiedon).toDateString() === today;
        }).length;

        const updatesEl = getElement('#recentUpdates');
        if (updatesEl) updatesEl.textContent = updatedToday;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // MODAL OPERATIONS - POWER PAGES OPTIMIZED
    // ═══════════════════════════════════════════════════════════════════════
    function openAddCustomerModal() {
        state.editingCustomerId = null;
        state.currentStep = 1;

        const modalTitle = getElement('#modalTitle span');
        if (modalTitle) modalTitle.textContent = 'Add New Customer';

        const form = getElement('#customerForm');
        if (form) form.reset();

        const countryEl = getElement('#country');
        if (countryEl && countryEl.value) {
            handleCountryChange({ target: countryEl });
        }

        const countryCodeInput = getElement('#country1');
        if (countryCodeInput) {
            countryCodeInput.value = '';
            countryCodeInput.disabled = false;
            delete countryCodeInput.dataset.suggested;
        }


        const codeInput = getElement('#customerCode');
        if (codeInput) codeInput.disabled = false;

        const codeValidation = getElement('#codeValidation');
        if (codeValidation) {
            codeValidation.className = 'validation-message-cm';
            codeValidation.style.display = 'none';
            codeValidation.innerHTML = '';
        }

        // Hide all custom "Other" inputs
        ['#countryCustom', '#paymentTermsCustom', '#cobCustom', '#salesRepCustom'].forEach(selector => {
            const el = getElement(selector);
            if (el) {
                el.style.display = 'none';
                el.required = false;
                el.value = '';
            }
        });

        // Reset Notify Party 2 visibility
        const notifyParty2Container = getElement('#notifyParty2Container');
        const addNotifyPartyBtn = getElement('#addNotifyPartyBtn');
        if (notifyParty2Container) notifyParty2Container.style.display = 'none';
        if (addNotifyPartyBtn) addNotifyPartyBtn.style.display = 'block';

        // Reset LP Comments checkbox and container
        const hasLpCommentsCheckbox = getElement('#hasLpComments');
        const lpCommentsContainer = getElement('#lpCommentsContainer');
        const lpCommentsTextarea = getElement('#lpComments');
        if (hasLpCommentsCheckbox) hasLpCommentsCheckbox.checked = false;
        if (lpCommentsContainer) lpCommentsContainer.style.display = 'none';
        if (lpCommentsTextarea) lpCommentsTextarea.value = '';

        // Reset Customer Models state and UI
        state.customerModels = [];
        state.pendingCustomerModels = [];
        state.editingModelId = null;
        state.editingPendingModelIndex = null;
        hideCustomerModelForm();
        renderCustomerModelsList();

        updateWizardUI();
        const modal = getElement('#customerModal');
        if (modal) {
            modal.style.display = 'flex';
            modal.classList.add('active');
            document.documentElement.style.overflow = 'hidden';
            document.body.style.overflow = 'hidden';
            log('Modal opened: Add mode');
        }
    }

    /**
     * ENHANCED: Handle editing with "Other" option support
     */
    async function handleEditCustomer(customerId) {
        log('handleEditCustomer called with ID:', customerId);
        state.editingCustomerId = customerId;
        state.currentStep = 1;

        const modalTitle = getElement('#modalTitle span');
        if (modalTitle) modalTitle.textContent = 'Edit Customer';

        showLoader(true);
        try {
            const customer = await getCustomerById(customerId);

            if (customer) {
                // Basic fields (removed old address fields - now using Customer Models)
                const fields = {
                    '#customerCode': customer.cr650_customercodes || '',
                    '#customerName': customer.cr650_customername || '',
                    '#destinationPort': customer.cr650_destinationport || '',
                    '#currency': customer.cr650_currency || '',
                    '#organizationId': customer.cr650_organizationid || '',
                    '#country1': customer.cr650_country_1 || '',
                    '#notifyParty1': customer.cr650_notifyparty1 || '',
                    '#notifyParty2': customer.cr650_notifyparty2 || '',
                    '#lpComments': customer.cr650_loadingplancomments || ''
                };

                // Fetch customer models for this customer
                await fetchCustomerModels(customerId);
                renderCustomerModelsList();

                // Handle LP Comments checkbox and textarea
                const hasLpCommentsCheckbox = getElement('#hasLpComments');
                const lpCommentsContainer = getElement('#lpCommentsContainer');
                const hasLpComments = customer.cr650_loadingplancomments && customer.cr650_loadingplancomments.trim() !== '';

                if (hasLpCommentsCheckbox) {
                    hasLpCommentsCheckbox.checked = hasLpComments;
                }
                if (lpCommentsContainer) {
                    lpCommentsContainer.style.display = hasLpComments ? 'block' : 'none';
                }

                // 🔁 SMART COUNTRY CODE HANDLING (EDIT MODE)
                const countryCodeInput = getElement('#country1');

                // Always allow editing in Edit mode
                countryCodeInput.disabled = false;

                // If ISO-2 exists, store it as a suggestion (for soft validation)
                if (customer.cr650_country_1 && customer.cr650_country_1.length === 2) {
                    countryCodeInput.dataset.suggested = customer.cr650_country_1.toUpperCase();
                } else {
                    delete countryCodeInput.dataset.suggested;
                }

                // Show Notify Party 2 if it has data
                if (customer.cr650_notifyparty2) {
                    showNotifyParty2();
                }

                Object.keys(fields).forEach(selector => {
                    const el = getElement(selector);
                    if (el) el.value = fields[selector];
                });

                // Handle Country (check if it's in the dropdown or custom)
                const countryEl = getElement('#country');
                const countryCustomEl = getElement('#countryCustom');
                const countryValue = customer.cr650_country || '';

                // Check if country exists in dropdown options
                const countryOptions = Array.from(countryEl.options).map(opt => opt.value);
                if (countryOptions.includes(countryValue)) {
                    countryEl.value = countryValue;
                    handleCountryChange({ target: countryEl });
                    if (countryCustomEl) {
                        countryCustomEl.style.display = 'none';
                        countryCustomEl.required = false;
                    }
                } else if (countryValue) {
                    // Custom country
                    countryEl.value = 'Other';
                    if (countryCustomEl) {
                        countryCustomEl.value = countryValue;
                        countryCustomEl.style.display = 'block';
                        countryCustomEl.required = true;
                    }
                }

                // Handle Payment Terms
                const paymentTermsEl = getElement('#paymentTerms');
                const paymentTermsCustomEl = getElement('#paymentTermsCustom');
                const paymentTermsValue = customer.cr650_paymentterms || '';

                const paymentOptions = Array.from(paymentTermsEl.options).map(opt => opt.value);
                if (paymentOptions.includes(paymentTermsValue)) {
                    paymentTermsEl.value = paymentTermsValue;
                    if (paymentTermsCustomEl) {
                        paymentTermsCustomEl.style.display = 'none';
                        paymentTermsCustomEl.required = false;
                    }
                } else if (paymentTermsValue) {
                    paymentTermsEl.value = 'Other';
                    if (paymentTermsCustomEl) {
                        paymentTermsCustomEl.value = paymentTermsValue;
                        paymentTermsCustomEl.style.display = 'block';
                        paymentTermsCustomEl.required = true;
                    }
                }

                // Handle COB
                const cobEl = getElement('#cob');
                const cobCustomEl = getElement('#cobCustom');
                const cobValue = customer.cr650_cob || '';

                const cobOptions = Array.from(cobEl.options).map(opt => opt.value);
                if (cobOptions.includes(cobValue)) {
                    cobEl.value = cobValue;
                    if (cobCustomEl) {
                        cobCustomEl.style.display = 'none';
                        cobCustomEl.required = false;
                    }
                } else if (cobValue) {
                    cobEl.value = 'Other';
                    if (cobCustomEl) {
                        cobCustomEl.value = cobValue;
                        cobCustomEl.style.display = 'block';
                        cobCustomEl.required = true;
                    }
                }

                // Handle Sales Rep (check if it's in the dropdown or custom)
                const salesRepEl = getElement('#salesRep');
                const salesRepCustomEl = getElement('#salesRepCustom');
                const salesRepValue = customer.cr650_salesrepresentativename || '';

                // Check if sales rep exists in dropdown options
                const salesRepOptions = Array.from(salesRepEl.options).map(opt => opt.value);
                if (salesRepOptions.includes(salesRepValue)) {
                    salesRepEl.value = salesRepValue;
                    if (salesRepCustomEl) {
                        salesRepCustomEl.style.display = 'none';
                        salesRepCustomEl.required = false;
                    }
                } else if (salesRepValue) {
                    // Custom sales rep
                    salesRepEl.value = 'Other';
                    if (salesRepCustomEl) {
                        salesRepCustomEl.value = salesRepValue;
                        salesRepCustomEl.style.display = 'block';
                        salesRepCustomEl.required = true;
                    }
                }

                const codeInput = getElement('#customerCode');
                if (codeInput) codeInput.disabled = true;

                updateWizardUI();

                const modal = getElement('#customerModal');
                if (modal) {
                    modal.style.display = 'flex';
                    modal.classList.add('active');
                    document.documentElement.style.overflow = 'hidden';
                    document.body.style.overflow = 'hidden';
                    log('Modal opened: Edit mode');
                }
            }
        } catch (error) {
            showToast('Failed to load customer details', 'error');
        } finally {
            showLoader(false);
        }
    }


    /**
     * Handle deleting a customer with confirmation
     */
    async function handleDeleteCustomer(customerId) {
        log('handleDeleteCustomer called with ID:', customerId);

        // Find customer name for confirmation message
        const customer = state.customers.find(c => c.cr650_updated_dcl_customerid === customerId);
        const customerName = customer ? customer.cr650_customername : 'this customer';
        const customerCode = customer ? customer.cr650_customercodes : '';

        // Confirm deletion
        if (!confirm(`Are you sure you want to delete "${customerName}" (${customerCode})?\n\nThis action cannot be undone.`)) {
            return;
        }

        showLoader(true);
        try {
            // Delete via Web API with anti-forgery token
            await safeAjax({
                type: 'DELETE',
                url: `${CONFIG.apiPath}/${CONFIG.entitySetName}(${customerId})`
            });

            showToast('Customer deleted successfully!', 'success');
            await loadCustomers();

        } catch (error) {
            console.error('Error deleting customer:', error);
            showToast(
                error?.responseText || error?.statusText || 'Failed to delete customer',
                'error'
            );
        } finally {
            showLoader(false);
        }
    }
    function closeModal() {
        const modal = getElement('#customerModal');
        if (modal) {
            modal.classList.remove('active');
            modal.style.display = 'none';
            log('Modal closed');
        }

        const countryCodeInput = getElement('#country1');
        if (countryCodeInput) {
            delete countryCodeInput.dataset.suggested;
        }


        document.documentElement.style.overflow = 'auto';
        document.body.style.overflow = 'auto';
        state.editingCustomerId = null;
        state.currentStep = 1;
        state.pendingCustomerModels = [];
        state.editingPendingModelIndex = null;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SEARCH & FILTER
    // ═══════════════════════════════════════════════════════════════════════
    function handleSearch(e) {
        const searchTerm = e.target.value.toLowerCase().trim();
        const countryFilterEl = getElement('#countryFilter');
        const countryValue = countryFilterEl ? countryFilterEl.value : '';

        state.filteredCustomers = state.customers.filter(customer => {
            const matchesSearch = !searchTerm ||
                (customer.cr650_customercodes || '').toLowerCase().includes(searchTerm) ||
                (customer.cr650_customername || '').toLowerCase().includes(searchTerm) ||
                (customer.cr650_country || '').toLowerCase().includes(searchTerm);

            const matchesCountry = !countryValue || customer.cr650_country === countryValue;

            return matchesSearch && matchesCountry;
        });

        state.currentPage = 1;
        renderCustomerTable();
    }

    function handleFilter(e) {
        const searchInput = getElement('#searchInput');
        if (searchInput) {
            handleSearch({ target: searchInput });
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // PAGINATION
    // ═══════════════════════════════════════════════════════════════════════
    function changePage(direction) {
        const newPage = state.currentPage + direction;
        if (newPage >= 1 && newPage <= state.totalPages) {
            state.currentPage = newPage;
            renderCustomerTable();

            const tableCard = document.querySelector('.table-card-cm');
            if (tableCard) {
                tableCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // NOTIFY PARTY CONTROLS
    // ═══════════════════════════════════════════════════════════════════════
    function showNotifyParty2() {
        const container = getElement('#notifyParty2Container');
        const btnContainer = getElement('#addNotifyPartyBtn');
        
        if (container) {
            container.style.display = 'block';
        }
        if (btnContainer) {
            btnContainer.style.display = 'none';
        }
    }

    function hideNotifyParty2() {
        const container = getElement('#notifyParty2Container');
        const btnContainer = getElement('#addNotifyPartyBtn');
        const notifyParty2Field = getElement('#notifyParty2');
        
        if (container) {
            container.style.display = 'none';
        }
        if (btnContainer) {
            btnContainer.style.display = 'block';
        }
        if (notifyParty2Field) {
            notifyParty2Field.value = ''; // Clear the field when removing
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // EXPORT TO EXCEL
    // ═══════════════════════════════════════════════════════════════════════
    function exportToExcel() {
        try {
            // Check if ExcelJS is available
            if (typeof ExcelJS === 'undefined') {
                showToast('Excel library not loaded. Please refresh the page.', 'error');
                return;
            }

            // ═══════════════════════════════════════════════════════════════
            // CREATE WORKBOOK & WORKSHEET
            // ═══════════════════════════════════════════════════════════════
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Customer Database', {
                properties: { tabColor: { argb: 'FF006633' } }
            });

            // ═══════════════════════════════════════════════════════════════
            // PREMIUM COLOR PALETTE
            // ═══════════════════════════════════════════════════════════════
            const colors = {
                headerPrimary: 'FF004d26',      // Dark green
                headerAccent: 'FFF7A900',       // Yellow
                headerText: 'FFFFFFFF',         // White
                
                titleBg: 'FFF0F9F4',           // Light green
                titleText: 'FF004d26',         // Dark green
                
                primaryGreen: 'FF006633',
                primaryGreenLight: 'FF00803f',
                
                white: 'FFFFFFFF',
                gray50: 'FFF9FAFB',
                gray100: 'FFF3F4F6',
                gray700: 'FF374151',
                gray800: 'FF1F2937',
                
                borderDark: 'FF006633',
                borderLight: 'FFD1D5DB',
                borderMedium: 'FF9CA3AF'
            };

            // ═══════════════════════════════════════════════════════════════
            // SET COLUMN WIDTHS
            // ═══════════════════════════════════════════════════════════════
            worksheet.columns = [
                { width: 16 },  // Customer Code
                { width: 38 },  // Customer Name
                { width: 12 },  // Currency
                { width: 22 },  // Payment Terms
                { width: 14 },  // Country Code
                { width: 18 },  // Country
                { width: 40 },  // Notify Party 1
                { width: 40 },  // Notify Party 2
                { width: 20 },  // Organization ID
                { width: 18 },  // COB
                { width: 22 },  // Destination Port
                { width: 28 }   // Sales Representative
            ];

            // ═══════════════════════════════════════════════════════════════
            // ROW 1: TOP SPACING
            // ═══════════════════════════════════════════════════════════════
            worksheet.addRow([]);
            worksheet.getRow(1).height = 10;

            // ═══════════════════════════════════════════════════════════════
            // ROW 2: MAIN TITLE
            // ═══════════════════════════════════════════════════════════════
            const titleRow = worksheet.addRow(['CUSTOMER MANAGEMENT SYSTEM']);
            titleRow.height = 40;
            worksheet.mergeCells('A2:L2');
            
            const titleCell = worksheet.getCell('A2');
            titleCell.font = {
                name: 'Segoe UI',
                size: 20,
                bold: true,
                color: { argb: colors.titleText }
            };
            titleCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colors.titleBg }
            };
            titleCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            titleCell.border = {
                top: { style: 'medium', color: { argb: colors.borderDark } },
                left: { style: 'medium', color: { argb: colors.borderDark } },
                right: { style: 'medium', color: { argb: colors.borderDark } }
            };

            // ═══════════════════════════════════════════════════════════════
            // ROW 3: SUBTITLE
            // ═══════════════════════════════════════════════════════════════
            const subtitleRow = worksheet.addRow(['Complete Customer Database Export']);
            subtitleRow.height = 25;
            worksheet.mergeCells('A3:L3');
            
            const subtitleCell = worksheet.getCell('A3');
            subtitleCell.font = {
                name: 'Segoe UI',
                size: 14,
                italic: true,
                color: { argb: colors.gray700 }
            };
            subtitleCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colors.titleBg }
            };
            subtitleCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            subtitleCell.border = {
                left: { style: 'medium', color: { argb: colors.borderDark } },
                right: { style: 'medium', color: { argb: colors.borderDark } }
            };

            // ═══════════════════════════════════════════════════════════════
            // ROW 4: TIMESTAMP & INFO
            // ═══════════════════════════════════════════════════════════════
            const timestamp = new Date().toLocaleString('en-US', {
                weekday: 'long',
                year: 'numeric',
                month: 'long',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
            const infoText = `Exported: ${timestamp} • Total Records: ${state.filteredCustomers.length}`;
            
            const infoRow = worksheet.addRow([infoText]);
            infoRow.height = 20;
            worksheet.mergeCells('A4:L4');
            
            const infoCell = worksheet.getCell('A4');
            infoCell.font = {
                name: 'Segoe UI',
                size: 10,
                italic: true,
                color: { argb: colors.gray700 }
            };
            infoCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colors.titleBg }
            };
            infoCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            infoCell.border = {
                bottom: { style: 'medium', color: { argb: colors.borderDark } },
                left: { style: 'medium', color: { argb: colors.borderDark } },
                right: { style: 'medium', color: { argb: colors.borderDark } }
            };

            // ═══════════════════════════════════════════════════════════════
            // ROWS 5-6: SPACING
            // ═══════════════════════════════════════════════════════════════
            worksheet.addRow([]);
            worksheet.getRow(5).height = 8;
            worksheet.addRow([]);
            worksheet.getRow(6).height = 8;

            // ═══════════════════════════════════════════════════════════════
            // ROW 7: COLUMN HEADERS
            // ═══════════════════════════════════════════════════════════════
            const headers = [
                'CUSTOMER CODE',
                'CUSTOMER NAME',
                'CURRENCY',
                'PAYMENT TERMS',
                'COUNTRY CODE',
                'COUNTRY',
                'NOTIFY PARTY 1',
                'NOTIFY PARTY 2',
                'ORGANIZATION ID',
                'COB',
                'DESTINATION PORT',
                'SALES REPRESENTATIVE'
            ];
            const headerRow = worksheet.addRow(headers);
            headerRow.height = 40;
            
            headerRow.eachCell((cell) => {
                cell.font = {
                    name: 'Segoe UI',
                    size: 11,
                    bold: true,
                    color: { argb: colors.headerText }
                };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: colors.headerPrimary }
                };
                cell.alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                    wrapText: true
                };
                cell.border = {
                    top: { style: 'medium', color: { argb: colors.borderDark } },
                    bottom: { style: 'medium', color: { argb: colors.headerAccent } },
                    left: { style: 'thin', color: { argb: colors.primaryGreenLight } },
                    right: { style: 'thin', color: { argb: colors.primaryGreenLight } }
                };
            });

            // ═══════════════════════════════════════════════════════════════
            // DATA ROWS
            // ═══════════════════════════════════════════════════════════════
            state.filteredCustomers.forEach((customer, index) => {
                const rowData = [
                    customer.cr650_customercodes || '',
                    customer.cr650_customername || '',
                    customer.cr650_currency || '',
                    customer.cr650_paymentterms || '',
                    customer.cr650_country_1 || '',
                    customer.cr650_country || '',
                    customer.cr650_notifyparty1 || '',
                    customer.cr650_notifyparty2 || '',
                    customer.cr650_organizationid || '',
                    customer.cr650_cob || '',
                    customer.cr650_destinationport || '',
                    customer.cr650_salesrepresentativename || ''
                ];
                
                const dataRow = worksheet.addRow(rowData);
                dataRow.height = 85;
                
                const isAlternating = index % 2 === 1;
                const bgColor = isAlternating ? colors.gray50 : colors.white;
                
                dataRow.eachCell((cell, colNumber) => {
                    // Customer Code styling (bold green)
                    if (colNumber === 1) {
                        cell.font = {
                            name: 'Segoe UI',
                            size: 10,
                            bold: true,
                            color: { argb: colors.primaryGreen }
                        };
                    } else {
                        cell.font = {
                            name: 'Segoe UI',
                            size: 10,
                            color: { argb: colors.gray800 }
                        };
                    }
                    
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: bgColor }
                    };
                    
                    cell.alignment = {
                        vertical: 'top',
                        horizontal: 'left',
                        wrapText: true
                    };
                    
                    cell.border = {
                        left: { style: 'thin', color: { argb: colors.borderLight } },
                        right: { style: 'thin', color: { argb: colors.borderLight } },
                        bottom: { style: 'hair', color: { argb: colors.borderLight } }
                    };
                });
            });

            // ═══════════════════════════════════════════════════════════════
            // FOOTER ROWS
            // ═══════════════════════════════════════════════════════════════
            worksheet.addRow([]);
            worksheet.addRow([]);
            
            // Copyright row
            const copyrightRow = worksheet.addRow(['© 2026 Customer Management System • Technolube International']);
            copyrightRow.height = 25;
            const copyrightRowNum = copyrightRow.number;
            worksheet.mergeCells(`A${copyrightRowNum}:L${copyrightRowNum}`);
            
            const copyrightCell = worksheet.getCell(`A${copyrightRowNum}`);
            copyrightCell.font = {
                name: 'Segoe UI',
                size: 9,
                bold: true,
                color: { argb: colors.gray700 }
            };
            copyrightCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colors.gray100 }
            };
            copyrightCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            copyrightCell.border = {
                top: { style: 'thin', color: { argb: colors.borderMedium } }
            };
            
            // Confidential row
            const confidentialRow = worksheet.addRow(['Confidential Business Information - For Internal Use Only']);
            confidentialRow.height = 20;
            const confidentialRowNum = confidentialRow.number;
            worksheet.mergeCells(`A${confidentialRowNum}:L${confidentialRowNum}`);
            
            const confidentialCell = worksheet.getCell(`A${confidentialRowNum}`);
            confidentialCell.font = {
                name: 'Segoe UI',
                size: 8,
                italic: true,
                color: { argb: colors.gray700 }
            };
            confidentialCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colors.gray100 }
            };
            confidentialCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            confidentialCell.border = {
                bottom: { style: 'medium', color: { argb: colors.borderDark } },
                left: { style: 'medium', color: { argb: colors.borderDark } },
                right: { style: 'medium', color: { argb: colors.borderDark } }
            };

            // ═══════════════════════════════════════════════════════════════
            // FREEZE PANES
            // ═══════════════════════════════════════════════════════════════
            worksheet.views = [
                { state: 'frozen', xSplit: 0, ySplit: 7 }
            ];

            // ═══════════════════════════════════════════════════════════════
            // WORKBOOK PROPERTIES
            // ═══════════════════════════════════════════════════════════════
            workbook.creator = 'Customer Management System';
            workbook.lastModifiedBy = 'Technolube';
            workbook.created = new Date();
            workbook.modified = new Date();

            // ═══════════════════════════════════════════════════════════════
            // GENERATE FILENAME
            // ═══════════════════════════════════════════════════════════════
            const now = new Date();
            const dateStr = now.toISOString().split('T')[0];
            const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-');
            const filename = `CustomerDatabase_${dateStr}_${timeStr}.xlsx`;

            // ═══════════════════════════════════════════════════════════════
            // EXPORT FILE
            // ═══════════════════════════════════════════════════════════════
            workbook.xlsx.writeBuffer().then(function(buffer) {
                const blob = new Blob([buffer], { 
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                });
                const url = window.URL.createObjectURL(blob);
                const anchor = document.createElement('a');
                anchor.href = url;
                anchor.download = filename;
                anchor.click();
                window.URL.revokeObjectURL(url);
                
                showToast(`✓ Exported ${state.filteredCustomers.length} customers successfully!`, 'success');
            });

        } catch (error) {
            console.error('Excel export error:', error);
            showToast('Failed to export Excel file: ' + error.message, 'error');
        }
    }
    // ═══════════════════════════════════════════════════════════════════════
    // AUTO-INITIALIZE ON DOM READY
    // ═══════════════════════════════════════════════════════════════════════
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize);
    } else {
        initialize();
    }

    // ═══════════════════════════════════════════════════════════════════════
    // APPLICATION READY INDICATOR
    // ═══════════════════════════════════════════════════════════════════════
    console.log('%c✅ Customer Management System v5.0 (Customer Models Feature) Ready!', 'color: #006633; font-size: 14px; font-weight: bold;');

})();