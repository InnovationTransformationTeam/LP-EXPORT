(function (w, d) {
    "use strict";

    /* ============================================================
       0) CONFIG / CONSTANTS
       ============================================================ */

    const DCL_API = "/_api/cr650_dcl_masters";
    const CUSTOMER_API = "/_api/cr650_updated_dcl_customers";
    const OUTSTANDING_API = "/_api/cr650_dcl_outstanding_reports";

    const NOTIFY_API = "/_api/cr650_dcl_notify_parties";
    const BRANDS_API = "/_api/cr650_dcl_brands";
    const DCL_ORDERS_API = "/_api/cr650_dcl_orders";
    const SHIPPING_LINES_API = "/_api/cr650_dcl_shippinglines";
    const DCL_DOCUMENTS_API = "/_api/cr650_dcl_documents";
    const CUSTOMER_MODEL_API = "/_api/cr650_dcl_customer_models";
    const BRANDS_LIST_API = "/_api/cr650_dcl_brands_lists";

    const CUSTOMER_FIELDS = [
        "cr650_updated_dcl_customerid",
        "cr650_customername",
        "cr650_customercodes",
        "cr650_country",
        "cr650_country_1",
        "cr650_cob",
        "cr650_paymentterms",
        "cr650_salesrepresentativename",
        "cr650_destinationport",
        "cr650_primaryaddress",
        "cr650_billto",
        "cr650_shipto",
        "cr650_consignee",
        "cr650_organizationid",
        "cr650_notifyparty1",
        "cr650_notifyparty2"
    ];

    const ORDER_FIELDS = [
        "cr650_customer_number",
        "cr650_order_no",
        "cr650_source_order_number",
        "cr650_cust_po_number",
        "createdon"
    ];

    const DCL_ORDER_FIELDS = [
        "cr650_dcl_orderid",
        "cr650_order_number",
        "cr650_ci_number",
        "_cr650_dcl_number_value",
        "createdon"
    ];

    const NOTIFY_FIELDS = [
        "cr650_dcl_notify_partyid",
        "cr650_notify_party",
        "_cr650_dcl_number_value",
        "createdon"
    ];

    const BRAND_FIELDS = [
        "cr650_dcl_brandid",
        "cr650_brand",
        "createdon",
        "_cr650_dcl_number_value"
    ];

    const CUSTOMER_MODEL_FIELDS = [
        "cr650_dcl_customer_modelid",
        "cr650_model_name",
        "cr650_info",
        "_cr650_dcl_id_value",
        "_cr650_cutomer_id_value",
        "cr650_is_related_to_a_customer",
        "createdon"
    ];

    const CURRENCY_SOURCES = [
        "https://openexchangerates.org/api/currencies.json",
        "https://api.frankfurter.app/currencies"
    ];
    const CURRENCY_CACHE_KEY = "dcl_currencies_v2";
    const CURRENCY_CACHE_TTL_MS = 24 * 60 * 60 * 1000;

    // Power Automate Endpoints
    const POWER_AUTOMATE_EXTRACT_ENDPOINT = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a30fd12c4ca34488b589d3261b5a2eba/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=X3vQ-GwyonKHmo-ZImPVA45_eYWelJgKYQU_tR6Vm2Y";
    const POWER_AUTOMATE_UPLOAD_ENDPOINT = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/41a79329e87f400fa632ea4e374e8eb0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EjqKzb2Ezk_4WSJ6yxrLA61AjOOlwvy7Y9usBb66K94";

    // Document type mapping for ED/BL/PI/PO files
    const DOCUMENT_TYPE_MAP = {
        'edFile': { type: 'Export Declaration#', field: 'cr650_ednumber' },
        'blFile': { type: 'Bill of Loading#', field: 'cr650_blnumber' },
        'piFile': { type: 'Proforma Invoice#', field: 'cr650_pinumber' },
        'poFile': { type: 'Purchase Order', field: 'cr650_ponumber' },
        'customerPoFile': { type: 'Customer Purchase Order', field: 'cr650_po_customer_number' },
        'lcFile': { type: 'Letter of Credit', field: 'cr650_lc_number' }
    };

    const ALLOWED_FILE_EXTENSIONS = ['pdf', 'docx', 'xlsx', 'png', 'jpg', 'jpeg'];

    /* ============================================================
       1) STATE
       ============================================================ */
    const cache = {
        customers: [],
        selectedCustomer: null,
        currencies: [],
        shippingLines: [],
        brandsList: [],
        currentDcl: null,
        currentId: null,
        original: {},
        customerAvailableModels: []
    };

    const $ = (sel, root = d) => root.querySelector(sel);
    const $$ = (sel, root = d) => Array.from(root.querySelectorAll(sel));
    const el = {};

    /* ============================================================
       2) UTILITIES
       ============================================================ */
    function getQueryParam(name) {
        return new URL(w.location.href).searchParams.get(name);
    }

    function isGuid(s) {
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
            .test(String(s || "").trim());
    }

    function escapeHtml(s) {
        return String(s ?? "").replace(/[&<>"']/g, m => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[m]));
    }

    function asISO(dateVal) {
        if (!dateVal) return null;
        const dt = new Date(dateVal + "T00:00:00");
        if (isNaN(dt)) return null;
        return dt.toISOString();
    }

    function cleanPatchObject(obj) {
        const cleaned = {};
        Object.keys(obj).forEach(k => {
            const v = obj[k];
            if (v === "" || v === undefined) return;
            cleaned[k] = v;
        });
        return cleaned;
    }

    /* ============================================================
       3) SAFE AJAX + LOADING BAR
       ============================================================ */
    const TopLoader = (() => {
        const bar = d.getElementById("top-loader");
        const shellNode = d.getElementById("dclWizard");
        let refs = 0;

        function show() {
            if (bar) bar.classList.remove("hidden");
            if (shellNode) shellNode.classList.add("blur-while-loading");
        }
        function hide() {
            if (bar) bar.classList.add("hidden");
            if (shellNode) shellNode.classList.remove("blur-while-loading");
        }

        return {
            begin() { refs++; if (refs === 1) show(); },
            end() { refs = Math.max(0, refs - 1); if (refs === 0) hide(); },
            flash(ms = 400) { this.begin(); setTimeout(() => this.end(), ms); }
        };
    })();

    w.TopLoader = TopLoader;

    w.safeAjax = function safeAjax(options) {
        return new Promise((resolve, reject) => {
            options = options || {};
            options.type = options.type || "GET";
            options.headers = options.headers || {};
            options.headers.Accept = options.headers.Accept
                || "application/json;odata.metadata=minimal";
            options.contentType = options.contentType
                || "application/json; charset=utf-8";

            async function go(token) {
                if (token) options.headers["__RequestVerificationToken"] = token;

                const init = {
                    method: options.type,
                    headers: options.headers
                };
                if (options.data) {
                    init.body = options.data;
                }

                TopLoader.begin();
                try {
                    const res = await fetch(options.url, init);
                    const text = await res.text();

                    if (!res.ok) {
                        try { reject(JSON.parse(text)); }
                        catch { reject(new Error(text.slice(0, 800) || `HTTP ${res.status}`)); }
                        return;
                    }

                    if (!text) { resolve({}); return; }

                    try { resolve(JSON.parse(text)); }
                    catch { reject(new Error("Response was not valid JSON.")); }
                } catch (e) {
                    reject(e);
                } finally {
                    TopLoader.end();
                }
            }

            if (w.shell && typeof shell.getTokenDeferred === "function") {
                shell.getTokenDeferred()
                    .done(go)
                    .fail(() => reject({ message: "Token API unavailable (user not authenticated?)" }));
            } else {
                go(null);
            }
        });
    };

    w.addEventListener("beforeunload", () => {
        const bar = d.getElementById("top-loader");
        if (bar) bar.classList.remove("hidden");
    });

    function setBusy(isBusy, msg) {
        const main = $(".main-section");
        const title = $("#pageTitle");
        if (isBusy) {
            if (main) main.setAttribute("aria-busy", "true");
            if (title) {
                title.dataset.prev = title.textContent;
                title.textContent = msg || "Loading…";
            }
        } else {
            if (main) main.removeAttribute("aria-busy");
            if (title && title.dataset.prev) {
                title.textContent = title.dataset.prev;
            }
        }
    }

    /* ============================================================
       4) CURRENCIES
       ============================================================ */
    function fetchWithTimeout(url, ms = 8000) {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), ms);

        return fetch(url, { signal: ctrl.signal })
            .then(res =>
                res.json()
                    .then(data => ({ ok: res.ok, data }))
                    .catch(() => ({ ok: res.ok, data: null }))
            )
            .catch(() => ({ ok: false, data: null }))
            .finally(() => clearTimeout(t));
    }

    function currenciesFromCache() {
        try {
            const cached = JSON.parse(w.localStorage.getItem(CURRENCY_CACHE_KEY) || "null");
            if (
                cached &&
                Array.isArray(cached.items) &&
                Date.now() - cached.savedAt < CURRENCY_CACHE_TTL_MS
            ) return cached.items;
        } catch { }
        return null;
    }

    function saveCurrenciesToCache(list) {
        try {
            w.localStorage.setItem(
                CURRENCY_CACHE_KEY,
                JSON.stringify({ savedAt: Date.now(), items: list })
            );
        } catch { }
    }

    function normalizeCurrencyMap(obj) {
        if (!obj || typeof obj !== "object") return {};
        const out = {};
        Object.keys(obj).forEach(code => {
            if (code && code.length === 3) {
                out[code.toUpperCase()] = String(obj[code] || code);
            }
        });
        return out;
    }

    async function ensureCurrenciesLoaded() {
        const cached = currenciesFromCache();
        if (cached) { cache.currencies = cached; return; }

        const results = await Promise.allSettled(
            CURRENCY_SOURCES.map(u => fetchWithTimeout(u))
        );
        const dicts = results
            .filter(r => r.status === "fulfilled" && r.value.ok)
            .map(r => r.value.data)
            .filter(Boolean);

        let union = {};
        dicts.forEach(dct => {
            const m = normalizeCurrencyMap(dct);
            Object.keys(m).forEach(c => {
                if (!union[c]) union[c] = m[c];
            });
        });

        if (!Object.keys(union).length) {
            union = {
                USD: "United States Dollar",
                SAR: "Saudi Riyal",
                EUR: "Euro",
                AED: "UAE Dirham",
                GBP: "Pound Sterling"
            };
        }

        cache.currencies = Object.keys(union)
            .sort()
            .map(code => ({ code, name: union[code] }));

        saveCurrenciesToCache(cache.currencies);
    }

    function populateCurrencySelect(selectEl) {
        if (!selectEl) return;
        const preserve = selectEl.value;
        const opts = ['<option value="">Select Currency</option>']
            .concat(
                cache.currencies.map(
                    c => `<option value="${c.code}">${escapeHtml(`${c.code} - ${c.name}`)}</option>`
                )
            );
        selectEl.innerHTML = opts.join("");
        if (preserve) selectEl.value = preserve;
    }

    function populateAllCurrencySelects() {
        $$(".currency-select").forEach(populateCurrencySelect);
    }

    function tryDefaultCurrencyByCountry() {
        const curSel = $("#currency");
        if (!curSel || curSel.value) return;
        const country = ($("#country")?.value || "").toUpperCase();
        const prefer = (code) => {
            if (curSel.querySelector(`option[value='${code}']`)) {
                curSel.value = code;
            }
        };

        if (country.includes("SAUDI") || country === "SA") {
            prefer("SAR");
        } else if (
            country.includes("UAE") ||
            country.includes("EMIRATES") ||
            country.includes("UNITED ARAB") ||
            country === "AE"
        ) {
            prefer("AED");
        } else if (
            country.includes("UNITED STATES") ||
            country.includes("USA") ||
            country === "US"
        ) {
            prefer("USD");
        } else if (
            country.includes("EUROPE") ||
            country.includes("GERMANY") ||
            country === "DE" ||
            country === "EU"
        ) {
            prefer("EUR");
        }
    }

    /* ============================================================
       SHIPPING LINES CRUD
       ============================================================ */
    async function loadAllShippingLines() {
        const select = "$select=cr650_shipping_line,cr650_dcl_shippinglineid";
        const orderby = "$orderby=cr650_shipping_line asc";
        const url = `${SHIPPING_LINES_API}?${select}&${orderby}`;

        try {
            const data = await safeAjax({ type: "GET", url });
            const rows = (data && (data.value || data)) || [];
            return rows.map(r => ({
                id: r.cr650_dcl_shippinglineid || "",
                name: r.cr650_shipping_line || ""
            }));
        } catch (err) {
            console.error("Failed to load shipping lines", err);
            return [];
        }
    }

    async function createShippingLine(name) {
        if (!name || !name.trim()) {
            throw new Error("Shipping line name is required");
        }

        const bodyObj = {
            cr650_shipping_line: name.trim()
        };

        try {
            const data = await safeAjax({
                type: "POST",
                url: SHIPPING_LINES_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });

            return {
                id: data.cr650_dcl_shippinglineid || "",
                name: data.cr650_shipping_line || ""
            };
        } catch (err) {
            console.error("Failed to create shipping line", err);
            throw err;
        }
    }

    function populateShippingLineSelect(selectEl, lines) {
        if (!selectEl) return;

        const currentValue = selectEl.value;

        selectEl.innerHTML = '<option value="">Select Shipping Line</option>';

        lines.forEach(line => {
            const opt = document.createElement("option");
            opt.value = line.name;
            opt.textContent = line.name;
            opt.dataset.lineId = line.id;
            selectEl.appendChild(opt);
        });

        const addOpt = document.createElement("option");
        addOpt.value = "__add_new__";
        addOpt.textContent = "+ Add new shipping line";
        selectEl.appendChild(addOpt);

        if (currentValue && selectEl.querySelector(`option[value="${currentValue}"]`)) {
            selectEl.value = currentValue;
        }
    }

    /* ============================================================
       BRANDS LIST CRUD (Lookup Table)
       ============================================================ */
    async function loadAllBrandsList() {
        const select = "$select=cr650_brand_name,cr650_dcl_brands_listid";
        const orderby = "$orderby=cr650_brand_name asc";
        const url = `${BRANDS_LIST_API}?${select}&${orderby}`;

        try {
            const data = await safeAjax({ type: "GET", url });
            const rows = (data && (data.value || data)) || [];
            return rows.map(r => ({
                id: r.cr650_dcl_brands_listid || "",
                name: r.cr650_brand_name || ""
            }));
        } catch (err) {
            console.error("Failed to load brands list", err);
            return [];
        }
    }

    async function createBrandInList(name) {
        if (!name || !name.trim()) {
            throw new Error("Brand name is required");
        }

        const bodyObj = {
            cr650_brand_name: name.trim()
        };

        try {
            const data = await safeAjax({
                type: "POST",
                url: BRANDS_LIST_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });

            return {
                id: data.cr650_dcl_brands_listid || "",
                name: data.cr650_brand_name || ""
            };
        } catch (err) {
            console.error("Failed to create brand in list", err);
            throw err;
        }
    }

    function populateBrandSelect(selectEl, brands) {
        if (!selectEl) return;

        const currentValue = selectEl.value;

        selectEl.innerHTML = '<option value="">Select Brand</option>';

        brands.forEach(brand => {
            const opt = document.createElement("option");
            opt.value = brand.name;
            opt.textContent = brand.name;
            opt.dataset.brandId = brand.id;
            selectEl.appendChild(opt);
        });

        const addOpt = document.createElement("option");
        addOpt.value = "__add_new__";
        addOpt.textContent = "+ Add new brand";
        selectEl.appendChild(addOpt);

        if (currentValue && selectEl.querySelector(`option[value="${currentValue}"]`)) {
            selectEl.value = currentValue;
        }
    }

    /* ============================================================
       5) CUSTOMERS
       ============================================================ */
    function normalizeCustomer(r) {
        const g = (x) => (x == null ? "" : x);

        return {
            id: g(r.cr650_updated_dcl_customerid),
            name: g(r.cr650_customername),
            code: g(r.cr650_customercodes),
            country: g(r.cr650_country),
            countryCode: g(r.cr650_country_1),
            orgId: g(r.cr650_organizationid),
            cob: g(r.cr650_cob),
            paymentTerms: g(r.cr650_paymentterms),
            salesRep: g(r.cr650_salesrepresentativename),
            destinationPort: g(r.cr650_destinationport),
            primaryAddress: g(r.cr650_primaryaddress),
            billTo: g(r.cr650_billto),
            shipTo: g(r.cr650_shipto),
            consignee: g(r.cr650_consignee),
            notifyParty1: g(r.cr650_notifyparty1),
            notifyParty2: g(r.cr650_notifyparty2)
        };
    }

    function loadAllCustomers() {
        const select = "$select=" + CUSTOMER_FIELDS.join(",");
        const top = "$top=5000";
        const url = `${CUSTOMER_API}?${select}&${top}`;

        const all = [];

        function loop(nextUrl) {
            return safeAjax({ type: "GET", url: nextUrl }).then(data => {
                const rows = data?.value || [];
                rows.forEach(r => all.push(normalizeCustomer(r)));

                if (data["@odata.nextLink"]) {
                    return loop(data["@odata.nextLink"]);
                }

                all.sort((a, b) =>
                    (a.name || "").localeCompare(b.name || "", undefined, {
                        sensitivity: "base"
                    })
                );

                return all;
            });
        }

        return loop(url);
    }

    function populateCustomerSelect(customers) {
        const sel = el.customerPicker;
        if (!sel) return;

        cache.customers = customers.slice();
        sel.innerHTML = '<option value="">Select Customer</option>';

        customers.forEach((c, i) => {
            const label = c.code ? `${c.name} (${c.code})` : (c.name || "(Unnamed)");
            const opt = document.createElement("option");
            opt.value = c.id || "";
            opt.textContent = label;
            opt.dataset.idx = String(i);
            sel.appendChild(opt);
        });

        sel.disabled = false;
    }

    function trySelectOrInjectCustomer(name, code) {
        if (el.customerNameText) {
            el.customerNameText.value = code ? `${name} (${code})` : (name || "");
        }

        const sel = el.customerPicker;
        if (!sel) return;
        const match =
            [...sel.options].find(o => code && o.textContent.includes(`(${code})`)) ||
            [...sel.options].find(o => name && o.textContent.trim().startsWith(name));
        if (match) sel.value = match.value;
    }

    /* ============================================================
       6) OUTSTANDING ORDERS
       ============================================================ */
    function normalizeOrderRecord(r) {
        const s = (v) => (v == null ? "" : String(v).trim());

        return {
            orderValue: s(r.cr650_order_no),
            label: s(r.cr650_order_no),
            custPoNumber: s(r.cr650_cust_po_number)
        };
    }

    function buildOrderQueryUrl(customerNumber) {
        const select = "$select=" + encodeURIComponent(ORDER_FIELDS.join(","));

        const safeCN = String(customerNumber || "")
            .trim()
            .replace(/'/g, "''");

        const filter = "$filter=" + encodeURIComponent(
            `contains(cr650_customer_number,'${safeCN}')`
        );

        const orderby = "$orderby=" + encodeURIComponent("createdon desc");
        const top = "$top=5000";

        return `${OUTSTANDING_API}?${select}&${filter}&${orderby}&${top}`;
    }

    function loadOrdersByCustomerNumber(customerNumber) {
        if (!customerNumber) return Promise.resolve([]);

        const all = [];
        function loop(url) {
            return safeAjax({ type: "GET", url })
                .then(data => {
                    const rows = (data && (data.value || data)) || [];
                    rows.forEach(r => all.push(normalizeOrderRecord(r)));
                    if (data && data["@odata.nextLink"]) {
                        return loop(data["@odata.nextLink"]);
                    }
                    return all;
                });
        }

        return loop(buildOrderQueryUrl(customerNumber));
    }

    function dedupeOrders(list) {
        const seen = new Set();
        const out = [];
        for (const o of list) {
            const key = o.orderValue;
            if (!key) continue;
            if (!seen.has(key)) {
                seen.add(key);
                out.push(o);
            }
        }
        return out;
    }

    function setOrderSelectLoading(isLoading, msg = "Loading orders…") {
        const sel = d.getElementById("orderNumberSelect");
        if (!sel) return;
        if (isLoading) {
            sel.innerHTML = `<option value="">${msg}</option>`;
            sel.disabled = true;
        } else {
            sel.disabled = false;
        }
    }

    function populateOrderSelect(orders) {
        const sel = d.getElementById("orderNumberSelect");
        if (!sel) return;

        const deduped = dedupeOrders(orders);

        if (!deduped.length) {
            sel.innerHTML = '<option value="">No outstanding orders for this customer</option>';
            sel.disabled = true;
            return;
        }

        sel.disabled = false;
        sel.innerHTML = '<option value="">Select Order Number</option>';

        deduped.forEach(o => {
            const opt = d.createElement("option");
            opt.value = o.orderValue;
            opt.textContent = o.label;
            opt.dataset.custPoNumber = o.custPoNumber || "";
            sel.appendChild(opt);
        });
    }

    function clearOrderSelectAndTable() {
        const sel = d.getElementById("orderNumberSelect");
        const tbody = d.querySelector("#orderNumbersTable tbody");
        const tableSection = d.getElementById("orderNumbersTableSection");
        if (sel) {
            sel.innerHTML = '<option value="">Select Order Number</option>';
            sel.disabled = false;
        }
        if (tbody) tbody.innerHTML = "";
        if (tableSection) tableSection.style.display = "none";
    }

    function clearOrderTableOnly() {
        const tbody = d.querySelector("#orderNumbersTable tbody");
        const tableSection = d.getElementById("orderNumbersTableSection");

        if (tbody) tbody.innerHTML = "";
        if (tableSection) tableSection.style.display = "none";
    }

    async function reloadOrdersForCustomerNumber(newCustNum) {
        if (!newCustNum) {
            clearOrderSelectAndTable();
            return;
        }
        setOrderSelectLoading(true);
        try {
            const allOrders = await loadOrdersByCustomerNumber(newCustNum);
            populateOrderSelect(allOrders);
        } catch (errOrders) {
            console.error("Order lookup failed:", errOrders);
            const orderSel = d.getElementById("orderNumberSelect");
            if (orderSel) {
                orderSel.innerHTML = '<option value="">Failed to load orders</option>';
                orderSel.disabled = true;
            }
        } finally {
            setOrderSelectLoading(false);
        }
    }

    /* ============================================================
       6.5) DCL ORDERS CHILD TABLE
       ============================================================ */
    function buildDclOrdersBaseUrl() {
        const select = "$select=" + encodeURIComponent(DCL_ORDER_FIELDS.join(","));
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${DCL_ORDERS_API}?${select}&${orderby}`;
    }

    async function loadAllDclOrdersRaw() {
        const records = [];
        async function loop(nextUrl) {
            const data = await safeAjax({ type: "GET", url: nextUrl });
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }
        await loop(buildDclOrdersBaseUrl());
        return records;
    }

    async function loadDclOrdersForCurrent() {
        if (!cache.currentId) return [];
        const every = await loadAllDclOrdersRaw();
        return every.filter(r =>
            (r._cr650_dcl_number_value || "").toLowerCase() === cache.currentId.toLowerCase()
        );
    }

    async function createDclOrder(orderNumber, ciNumber = "") {
        if (!cache.currentId) {
            throw new Error("No DCL id in cache.currentId.");
        }

        const bodyObj = {
            cr650_order_number: orderNumber || "",
            cr650_ci_number: ciNumber || "",
            "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + cache.currentId + ")"
        };

        const data = await safeAjax({
            type: "POST",
            url: DCL_ORDERS_API,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            },
            data: JSON.stringify(bodyObj)
        });

        return data;
    }

    async function deleteDclOrder(dclOrderId) {
        if (!dclOrderId) return;
        const url = `${DCL_ORDERS_API}(${encodeURIComponent(dclOrderId)})`;
        await safeAjax({
            type: "DELETE",
            url,
            headers: { "If-Match": "*" }
        });
        return { ok: true };
    }

    async function deleteAllOrdersForCurrentDcl() {
        if (!cache.currentId) return;
        try {
            const list = await loadDclOrdersForCurrent();
            for (const row of list) {
                const oid = row.cr650_dcl_orderid;
                if (!oid) continue;
                try { await deleteDclOrder(oid); }
                catch (subErr) {
                    console.warn("Failed to delete dcl_order row", subErr);
                }
            }
        } catch (e) {
            console.error("Failed bulk delete dcl_orders for DCL", e);
        }
    }

    function renderExistingDclOrders(list) {
        const tableSection = $("#orderNumbersTableSection");
        const tbody = $("#orderNumbersTable tbody");
        if (!tbody || !tableSection) return;

        tbody.innerHTML = "";

        list.forEach(row => {
            const dclOrderId = row.cr650_dcl_orderid || "";
            const orderNumber = row.cr650_order_number || "";
            const ciNumber = row.cr650_ci_number || "";

            const tr = d.createElement("tr");
            tr.dataset.order = orderNumber;
            tr.dataset.dclOrderId = dclOrderId;

            const tdOrder = d.createElement("td");
            tdOrder.style.padding = "10px";
            tdOrder.textContent = orderNumber;

            const tdCI = d.createElement("td");
            tdCI.style.padding = "10px";
            tdCI.textContent = ciNumber || "";

            const tdAct = d.createElement("td");
            tdAct.style.padding = "10px";
            const rm = d.createElement("button");
            rm.type = "button";
            rm.textContent = "Remove";
            rm.className = "btn btn-outline";
            rm.addEventListener("click", async () => {
                const sure = confirm("Remove this order from this DCL?");
                if (!sure) return;

                try {
                    if (dclOrderId) {
                        await deleteDclOrder(dclOrderId);
                    }
                    tr.remove();
                    if (!tbody.children.length) {
                        tableSection.style.display = "none";
                    }
                } catch (errDelOne) {
                    console.error("Failed to delete dcl_order row", errDelOne);
                    alert("Couldn't remove this order from the DCL.");
                }
            });

            tdAct.appendChild(rm);

            tr.appendChild(tdOrder);
            tr.appendChild(tdCI);
            tr.appendChild(tdAct);

            tbody.appendChild(tr);
        });

        tableSection.style.display = list.length ? "" : "none";
    }

    /* ============================================================
       7) PARTY FIELD FILLER
       ============================================================ */
    function fillPartyUIFromRecord(partyType, record) {
        const billBox = $("#billToContainer");
        const billTA = $("#billTo");
        const mainTA = $("#shipTo");
        const partySel = $("#shipToLabel");

        if (partySel) {
            partySel.value = partyType;
        }

        const shipToVal = record?.cr650_shiptolocation || record?.shipTo || "";
        const billToVal = record?.cr650_billto || record?.billTo || "";
        const consigneeVal = record?.cr650_consignee || record?.consignee || "";
        const applicantVal = record?.cr650_applicant || record?.primaryAddress || "";

        if (partyType === "shipto") {
            if (mainTA) mainTA.value = shipToVal;
            if (billTA) billTA.value = billToVal;
            if (billBox) billBox.style.display = "";
            if (billTA) billTA.removeAttribute("disabled");
        } else if (partyType === "consignee") {
            if (mainTA) mainTA.value = consigneeVal;
            if (billBox) billBox.style.display = "none";
            if (billTA) billTA.setAttribute("disabled", "disabled");
        } else {
            if (mainTA) mainTA.value = applicantVal;
            if (billBox) billBox.style.display = "none";
            if (billTA) billTA.setAttribute("disabled", "disabled");
        }
    }

    /* ============================================================
       8) NOTIFY PARTY CRUD
       ============================================================ */
    function buildNotifyQueryUrl(dclGuid) {
        const select = "$select=" + encodeURIComponent(NOTIFY_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_dcl_number_value eq ${dclGuid}`);
        const orderby = "$orderby=" + encodeURIComponent("createdon desc");
        return `${NOTIFY_API}?${select}&${filter}&${orderby}`;
    }

    async function loadNotifyParties(dclGuid) {
        if (!dclGuid) return [];
        const records = [];

        async function loop(nextUrl) {
            const data = await safeAjax({ type: "GET", url: nextUrl });
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildNotifyQueryUrl(dclGuid));

        return records.map(normalizeNotifyParty);
    }

    function normalizeNotifyParty(r) {
        return {
            id: r.cr650_dcl_notify_partyid || "",
            text: r.cr650_notify_party || "",
            created: r.createdon || "",
            dclGuid: r._cr650_dcl_number_value || "",
            dclNumber: r["_cr650_dcl_number_value@OData.Community.Display.V1.FormattedValue"] || ""
        };
    }

    async function createNotifyParty(noteText) {
        if (!cache.currentId) throw new Error("No DCL id in cache.currentId.");

        const bodyObj = {
            cr650_notify_party: noteText,
            "cr650_Dcl_number@odata.bind": "/cr650_dcl_masters(" + cache.currentId + ")"
        };

        const data = await safeAjax({
            type: "POST",
            url: NOTIFY_API,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            },
            data: JSON.stringify(bodyObj)
        });

        return normalizeNotifyParty(data);
    }

    async function updateNotifyParty(id, noteText) {
        const url = `${NOTIFY_API}(${encodeURIComponent(id)})`;
        const bodyObj = { cr650_notify_party: noteText };

        await safeAjax({
            type: "PATCH",
            url,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "If-Match": "*"
            },
            data: JSON.stringify(bodyObj)
        });

        return { ok: true };
    }

    async function deleteNotifyParty(id) {
        const url = `${NOTIFY_API}(${encodeURIComponent(id)})`;
        await safeAjax({
            type: "DELETE",
            url,
            headers: { "If-Match": "*" }
        });
        return { ok: true };
    }

    function buildNotifyPartyRow(item, isNew) {
        const row = d.createElement("div");
        row.className = "notify-card";
        row.dataset.notifyId = item.id || "";

        const createdDate = item.created ? new Date(item.created) : null;
        const humanTime = (createdDate && !isNaN(createdDate))
            ? createdDate.toLocaleString()
            : "";

        row.innerHTML = `
            <textarea
                class="notify-text"
                rows="4"
                placeholder="Enter Notify Party details (name, address, contact)">${escapeHtml(item.text || "")}</textarea>

            <div class="notify-actions">
                <span class="notify-meta">${humanTime ? `Saved: ${humanTime}` : ""}</span>
                <div class="notify-buttons">
                    <button
                        type="button"
                        class="notify-save-btn"
                        data-mode="${isNew ? "create" : "update"}">
                        Save
                    </button>
                    <button
                        type="button"
                        class="notify-delete-btn ${isNew ? "hidden" : ""}">
                        Delete
                    </button>
                </div>
            </div>
        `;
        return row;
    }

    function renderNotifyPartyList(list) {
        const container = d.getElementById("notifyPartyContainer");
        if (!container) return;

        container.innerHTML = "";

        if (!list.length) {
            container.innerHTML =
                '<div style="font-size:13px; color:#6b7280; font-style:italic; margin-bottom:8px;">No notify parties added yet.</div>';
            return;
        }

        list.forEach(item => {
            const rowEl = buildNotifyPartyRow(item, false);
            container.appendChild(rowEl);
        });
    }

    function appendDraftNotifyPartyRow() {
        const container = d.getElementById("notifyPartyContainer");
        if (!container) return;

        if (container.querySelector('.notify-card[data-notify-id=""]')) return;

        const draftItem = {
            id: "",
            text: "",
            created: "",
            dclGuid: cache.currentId || "",
            dclNumber: ""
        };

        const rowEl = buildNotifyPartyRow(draftItem, true);
        container.appendChild(rowEl);
    }

    w.addNotifyPartyDraft = function () {
        appendDraftNotifyPartyRow();
    };

    /* ============================================================
       8b) CUSTOMER MODELS CRUD
       ============================================================ */
    function buildCustomerModelQueryUrl(dclGuid) {
        const select = "$select=" + encodeURIComponent(CUSTOMER_MODEL_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_dcl_id_value eq ${dclGuid}`);
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${CUSTOMER_MODEL_API}?${select}&${filter}&${orderby}`;
    }

    function normalizeCustomerModel(r) {
        return {
            id: r.cr650_dcl_customer_modelid || "",
            modelName: r.cr650_model_name || "",
            info: r.cr650_info || "",
            dclId: r._cr650_dcl_id_value || "",
            customerId: r._cr650_cutomer_id_value || "",
            isRelatedToCustomer: r.cr650_is_related_to_a_customer || false,
            created: r.createdon || ""
        };
    }

    async function loadCustomerModels(dclGuid) {
        if (!dclGuid) return [];
        const records = [];

        async function loop(nextUrl) {
            const data = await safeAjax({ type: "GET", url: nextUrl });
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildCustomerModelQueryUrl(dclGuid));
        return records.map(normalizeCustomerModel);
    }

    // Query for models linked to a specific customer (not DCL)
    function buildCustomerAvailableModelsQueryUrl(customerId) {
        const select = "$select=" + encodeURIComponent(CUSTOMER_MODEL_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_cutomer_id_value eq ${customerId} and cr650_is_related_to_a_customer eq true`);
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${CUSTOMER_MODEL_API}?${select}&${filter}&${orderby}`;
    }

    async function loadCustomerAvailableModels(customerId) {
        if (!customerId) return [];
        const records = [];

        async function loop(nextUrl) {
            const data = await safeAjax({ type: "GET", url: nextUrl });
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildCustomerAvailableModelsQueryUrl(customerId));
        return records.map(normalizeCustomerModel);
    }

    function renderCustomerAvailableModels(models) {
        const container = d.getElementById("customerAvailableModelsContainer");
        if (!container) return;

        if (!models || models.length === 0) {
            container.innerHTML = `
                <div style="text-align: center; padding: 16px; background: #f9f9f9; border: 1px dashed #ccc; border-radius: 8px;">
                    <i class="fas fa-inbox" style="font-size: 1.5rem; color: #999; margin-bottom: 8px;"></i>
                    <p style="color: #666; margin: 0; font-size: 0.9rem;">No customer models available.</p>
                    <p style="color: #999; font-size: 0.8rem; margin-top: 4px;">This customer has no saved models.</p>
                </div>
            `;
            return;
        }

        container.innerHTML = models.map((model, index) => `
            <div class="customer-available-model-card" data-model-id="${escapeHtml(model.id)}" data-index="${index}" style="
                background: #fff;
                border: 1px solid #e0e0e0;
                border-left: 4px solid var(--primary-green, #28a745);
                border-radius: 8px;
                padding: 12px 16px;
            ">
                <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                    <div style="flex: 1;">
                        <strong style="font-size: 1rem; color: #333;">${escapeHtml(model.modelName)}</strong>
                        <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #666; white-space: pre-wrap;">${model.info ? escapeHtml(model.info) : '<em style="color: #999;">No additional info</em>'}</p>
                    </div>
                    <div style="display: flex; gap: 6px; margin-left: 12px;">
                        <button type="button" class="btn-use-customer-model" data-model-name="${escapeHtml(model.modelName)}" data-model-info="${escapeHtml(model.info || '')}" style="
                            padding: 4px 10px;
                            background: #e8f5e9;
                            color: #2e7d32;
                            border: 1px solid #a5d6a7;
                            border-radius: 4px;
                            font-size: 0.75rem;
                            cursor: pointer;
                        "><i class="fas fa-copy"></i> Use</button>
                    </div>
                </div>
            </div>
        `).join('');

        // Add click handlers for "Use" buttons
        container.querySelectorAll('.btn-use-customer-model').forEach(btn => {
            btn.addEventListener('click', function() {
                const modelName = this.dataset.modelName || '';
                const modelInfo = this.dataset.modelInfo || '';
                // Add a draft row with pre-filled data
                appendDraftCustomerModelRow(modelName, modelInfo);
            });
        });
    }

    async function createCustomerModel(modelName, info) {
        if (!cache.currentId) throw new Error("No DCL id in cache.currentId.");

        const bodyObj = {
            cr650_model_name: modelName,
            cr650_info: info || "",
            cr650_is_related_to_a_customer: false,
            "cr650_dcl_id@odata.bind": "/cr650_dcl_masters(" + cache.currentId + ")"
        };

        const data = await safeAjax({
            type: "POST",
            url: CUSTOMER_MODEL_API,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            },
            data: JSON.stringify(bodyObj)
        });

        return normalizeCustomerModel(data);
    }

    async function updateCustomerModel(id, modelName, info) {
        const url = `${CUSTOMER_MODEL_API}(${encodeURIComponent(id)})`;
        const bodyObj = {
            cr650_model_name: modelName,
            cr650_info: info || ""
        };

        await safeAjax({
            type: "PATCH",
            url,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "If-Match": "*"
            },
            data: JSON.stringify(bodyObj)
        });

        return { ok: true };
    }

    async function deleteCustomerModel(id) {
        const url = `${CUSTOMER_MODEL_API}(${encodeURIComponent(id)})`;
        await safeAjax({
            type: "DELETE",
            url,
            headers: { "If-Match": "*" }
        });
        return { ok: true };
    }

    function buildCustomerModelRow(item, isNew) {
        const row = d.createElement("div");
        row.className = "customer-model-card";
        row.dataset.modelId = item.id || "";

        const createdDate = item.created ? new Date(item.created) : null;
        const humanTime = (createdDate && !isNaN(createdDate))
            ? createdDate.toLocaleString()
            : "";

        row.innerHTML = `
            <div class="model-field">
                <label class="model-label">Model Name *</label>
                <input
                    type="text"
                    class="model-name-input"
                    placeholder="Enter model name"
                    value="${escapeHtml(item.modelName || "")}"
                    required />
            </div>
            <div class="model-field">
                <label class="model-label">Model Information</label>
                <textarea
                    class="model-info-input"
                    rows="3"
                    placeholder="Enter additional information about this model">${escapeHtml(item.info || "")}</textarea>
            </div>
            <div class="model-actions">
                <span class="model-meta">${humanTime ? `Saved: ${humanTime}` : ""}</span>
                <div class="model-buttons">
                    <button
                        type="button"
                        class="model-save-btn"
                        data-mode="${isNew ? "create" : "update"}">
                        <i class="fas fa-save"></i> Save
                    </button>
                    <button
                        type="button"
                        class="model-delete-btn ${isNew ? "hidden" : ""}">
                        <i class="fas fa-trash"></i> Delete
                    </button>
                </div>
            </div>
        `;
        return row;
    }

    function renderCustomerModelList(list) {
        const container = d.getElementById("customerModelContainer");
        if (!container) return;

        container.innerHTML = "";

        if (!list.length) {
            container.innerHTML =
                '<div style="font-size:13px; color:#6b7280; font-style:italic; margin-bottom:8px;">No customer models added yet.</div>';
            return;
        }

        list.forEach(item => {
            const rowEl = buildCustomerModelRow(item, false);
            container.appendChild(rowEl);
        });
    }

    function appendDraftCustomerModelRow(prefillName = "", prefillInfo = "") {
        const container = d.getElementById("customerModelContainer");
        if (!container) return;

        // Remove "no models" placeholder if present
        const placeholder = container.querySelector('div[style*="font-style:italic"]');
        if (placeholder) placeholder.remove();

        const draftItem = {
            id: "",
            modelName: prefillName,
            info: prefillInfo,
            created: "",
            dclId: cache.currentId || "",
            customerId: "",
            isRelatedToCustomer: false
        };

        const rowEl = buildCustomerModelRow(draftItem, true);
        container.appendChild(rowEl);

        // Focus on the name input
        const nameInput = rowEl.querySelector(".model-name-input");
        if (nameInput) nameInput.focus();
    }

    w.addCustomerModelDraft = function (prefillName, prefillInfo) {
        appendDraftCustomerModelRow(prefillName, prefillInfo);
    };

    /* ============================================================
       9) BRANDS CRUD
       ============================================================ */
    function buildBrandsQueryUrl(dclGuid) {
        const select = "$select=" + encodeURIComponent(BRAND_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_dcl_number_value eq ${dclGuid}`);
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${BRANDS_API}?${select}&${filter}&${orderby}`;
    }

    function normalizeBrandRow(r) {
        return {
            id: r.cr650_dcl_brandid || "",
            name: r.cr650_brand || "",
            created: r.createdon || "",
            dclGuid: r._cr650_dcl_number_value || "",
            dclNumber: r["_cr650_dcl_number_value@OData.Community.Display.V1.FormattedValue"] || ""
        };
    }

    async function loadBrands(dclGuid) {
        if (!dclGuid) return [];
        const records = [];

        async function loop(nextUrl) {
            const data = await safeAjax({ type: "GET", url: nextUrl });
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildBrandsQueryUrl(dclGuid));

        return records.map(normalizeBrandRow);
    }

    async function createBrand(brandName) {
        if (!cache.currentId) {
            throw new Error("No DCL id in cache.currentId.");
        }

        const bodyObj = {
            cr650_brand: brandName,
            "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + cache.currentId + ")"
        };

        const data = await safeAjax({
            type: "POST",
            url: BRANDS_API,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            },
            data: JSON.stringify(bodyObj)
        });

        return normalizeBrandRow(data);
    }

    async function deleteBrand(id) {
        const url = `${BRANDS_API}(${encodeURIComponent(id)})`;

        await safeAjax({
            type: "DELETE",
            url,
            headers: {
                "If-Match": "*"
            }
        });

        return { ok: true };
    }

    function buildBrandChip(brandObj) {
        const chip = document.createElement("span");
        chip.className = "brand-chip";
        chip.dataset.value = brandObj.name;
        chip.dataset.brandId = brandObj.id || "";

        chip.innerHTML = `
            <span>${escapeHtml(brandObj.name)}</span>
            <button
                type="button"
                class="brand-remove-btn"
                aria-label="Remove brand"
            >×</button>
        `;
        return chip;
    }

    function renderBrandList(list) {
        const wrap = document.getElementById("brandList");
        if (!wrap) return;

        wrap.innerHTML = "";
        list.forEach(b => {
            const chip = buildBrandChip(b);
            wrap.appendChild(chip);
        });
    }

    /* ============================================================
       11) DCL MASTER LOAD / PATCH
       ============================================================ */
    function selectFields() {
        return [
            "cr650_dcl_masterid",
            "cr650_dclnumber",
            "createdon",
            "modifiedon",

            "cr650_customername",
            "cr650_customernumber",
            "cr650_company_type",
            "cr650_country",
            "cr650_organizationid",
            "cr650_cob",
            "cr650_countryoforigin",
            "cr650_paymentterms",
            "cr650_currencycode",
            "cr650_incoterms",

            "cr650_shiptolocation",
            "cr650_consignee",
            "cr650_billto",
            "cr650_applicant",
            "cr650_party",

            "cr650_ednumber",
            "cr650_blnumber",
            "cr650_pinumber",
            "cr650_ponumber",
            "cr650_po_customer_number",
            "cr650_lc_number",

            "cr650_actualreadinessdate",
            "cr650_lcissuedate",
            "cr650_lclatestshipmentdate",
            "cr650_banksubmissiondate",
            "cr650_lcexpirydate",

            "cr650_shippingline",
            "cr650_loadingport",
            "cr650_destinationport",
            "cr650_loadingdate",
            "cr650_sailing_date",

            "cr650_loading_type",
            "cr650_loadingtime",
            "cr650_loading_shift",

            "cr650_transportationmode",
            "cr650_salesrepresentativename",
            "cr650_description_goods_services",
            "cr650_status"
        ];
    }

    /* ============================================================
       DOCUMENT NUMBER EXTRACTION & UPLOAD VIA POWER AUTOMATE
       ============================================================ */

    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                const base64 = reader.result.split(',')[1];
                resolve(base64);
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    function getFileExtension(filename) {
        const parts = filename.split('.');
        return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
    }

    /**
     * Extract number from file via Power Automate
     */
    async function extractNumberFromFile(base64, extension) {
        const payload = {
            file_base64: base64,
            file_extension: extension
        };

        const response = await fetch(POWER_AUTOMATE_EXTRACT_ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        return await response.json();
    }

    /**
     * Upload file to SharePoint via Power Automate Flow
     */
    /**
  * Upload file to SharePoint via Power Automate Flow
  * @returns {Promise<{success: boolean, fileUrl: string}>}
  */
    async function uploadFileToSharePoint(masterId, fileBase64, docType, fileExtension) {
        if (!masterId || !fileBase64 || !docType) {
            throw new Error("Missing required parameters for file upload");
        }

        const payload = {
            id: masterId,
            fileContent: fileBase64,
            docType: docType,
            fileExtension: fileExtension || "pdf",
            language: "en"
        };

        const response = await fetch(POWER_AUTOMATE_UPLOAD_ENDPOINT, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        const text = await response.text();

        if (!response.ok) {
            throw new Error(`Upload Flow failed ${response.status}: ${text}`);
        }

        try {
            const result = JSON.parse(text);
            return {
                success: result.success || true,
                fileUrl: result.fileUrl || null,
                message: result.message || ""
            };
        } catch {
            return { success: true, fileUrl: null, message: text };
        }
    }

    /**
    * Create a document record in cr650_dcl_documents table
    */
    async function createDocumentRecord(masterId, docType, fileExtension, remarks = null, documentUrl = null) {
        if (!masterId || !docType) {
            throw new Error("Missing required parameters for document record creation");
        }

        const payload = {
            cr650_doc_type: docType,
            cr650_file_extention: fileExtension || "pdf",
            cr650_remarks: remarks,
            cr650_documenturl: documentUrl,  // ← Save the SharePoint URL
            // Leave currency and charge blank as requested
            "cr650_dcl_number@odata.bind": `/cr650_dcl_masters(${masterId})`
        };

        const response = await safeAjax({
            type: "POST",
            url: DCL_DOCUMENTS_API,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            },
            data: JSON.stringify(payload)
        });

        return response;
    }

    /**
     * Check if a document of this type already exists for this DCL
     */
    async function checkExistingDocument(masterId, docType) {
        if (!masterId || !docType) return null;

        const filter = `_cr650_dcl_number_value eq ${masterId} and cr650_doc_type eq '${docType.replace(/'/g, "''")}'`;
        const url = `${DCL_DOCUMENTS_API}?$filter=${encodeURIComponent(filter)}&$select=cr650_dcl_documentid,cr650_doc_type,cr650_documenturl`;

        try {
            const response = await safeAjax({ type: "GET", url });
            const records = response?.value || [];
            return records.length > 0 ? records[0] : null;
        } catch (err) {
            console.warn("Failed to check existing document:", err);
            return null;
        }
    }

    /**
     * Delete an existing document record
     */
    async function deleteDocumentRecord(documentId) {
        if (!documentId) return;

        const url = `${DCL_DOCUMENTS_API}(${documentId})`;
        await safeAjax({
            type: "DELETE",
            url,
            headers: { "If-Match": "*" }
        });
    }

    /**
     * Show visual success feedback after upload
     */
    function showUploadSuccess(fileInput, docType) {
        const successBadge = document.createElement('span');
        successBadge.className = 'upload-success-badge';
        successBadge.innerHTML = `✓ ${docType} uploaded`;
        successBadge.style.cssText = `
            display: inline-block;
            background: linear-gradient(135deg, #10b981 0%, #059669 100%);
            color: white;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 8px;
            animation: fadeIn 0.3s ease;
        `;

        const existingBadge = fileInput.parentNode?.querySelector('.upload-success-badge');
        if (existingBadge) existingBadge.remove();

        if (fileInput.parentNode) {
            fileInput.parentNode.appendChild(successBadge);

            setTimeout(() => {
                successBadge.style.opacity = '0';
                successBadge.style.transition = 'opacity 0.3s ease';
                setTimeout(() => successBadge.remove(), 300);
            }, 5000);
        }
    }

    /**
     * Create number dropdown that also handles file storage
     */
    function createNumberDropdownWithStorage(numbers, targetInputId, fileInput, base64, fileExtension, docConfig) {
        const targetInput = document.getElementById(targetInputId);
        if (!targetInput) return;

        const existingDropdown = document.getElementById(`${targetInputId}_dropdown`);
        if (existingDropdown) existingDropdown.remove();

        const select = document.createElement('select');
        select.id = `${targetInputId}_dropdown`;
        select.style.cssText = `
        padding: 0.875rem 1rem;
        border-radius: 12px;
        font-size: 0.95rem;
        margin-top: 8px;
        width: 100%;
        border: 1px solid #d1d5db;
        background-color: #fff;
    `;

        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'Select extracted number...';
        select.appendChild(defaultOption);

        numbers.forEach((item, index) => {
            const option = document.createElement('option');
            const numberValue = typeof item === 'string'
                ? item
                : (item.item || item.number || item.value || '');

            option.value = numberValue;
            option.textContent = numberValue || `Unknown Item ${index + 1}`;
            select.appendChild(option);
        });

        select.addEventListener('change', async function () {
            const selectedValue = this.value;

            if (selectedValue) {
                targetInput.value = selectedValue;
                targetInput.disabled = true;

                try {
                    // Upload to SharePoint
                    targetInput.value = 'Uploading to SharePoint...';

                    const uploadResult = await uploadFileToSharePoint(
                        cache.currentId,
                        base64,
                        docConfig.type,
                        fileExtension
                    );

                    // Get the SharePoint URL from the response
                    const sharePointUrl = uploadResult.fileUrl || null;

                    // Handle document record
                    targetInput.value = 'Saving document record...';

                    const existingDoc = await checkExistingDocument(cache.currentId, docConfig.type);
                    if (existingDoc) {
                        await deleteDocumentRecord(existingDoc.cr650_dcl_documentid);
                    }

                    const file = fileInput.files[0];
                    await createDocumentRecord(
                        cache.currentId,
                        docConfig.type,
                        fileExtension,
                        `Uploaded: ${file?.name || 'document'}`,
                        sharePointUrl  // ← Pass the SharePoint URL
                    );

                    // Update text field and save to DCL
                    targetInput.value = selectedValue;

                    if (cache.currentId) {
                        await patchSingleField(docConfig.field, selectedValue);
                    }

                    // Remove dropdown after successful save
                    this.remove();

                    showUploadSuccess(fileInput, docConfig.type);

                    // Log success with URL
                    console.log(`✅ Document uploaded successfully. SharePoint URL: ${sharePointUrl || 'N/A'}`);

                } catch (error) {
                    console.error('Error after number selection:', error);
                    alert(`Failed to save document: ${error.message}`);
                    targetInput.value = '';
                } finally {
                    targetInput.disabled = false;
                }
            }
        });

        if (fileInput.parentNode) {
            fileInput.parentNode.insertBefore(select, fileInput.nextSibling);
        } else {
            targetInput.parentNode.insertBefore(select, targetInput.nextSibling);
        }
    }

    /**
     * Enhanced document file upload handler
     * - Extracts number from document
     * - Uploads file to SharePoint via Flow
     * - Creates/updates record in cr650_dcl_documents
     */

    async function handleDocumentFileUploadWithStorage(fileInput, textInputId) {
        const file = fileInput.files[0];
        if (!file) return;

        const textInput = document.getElementById(textInputId);
        if (!textInput) {
            console.error(`Text input ${textInputId} not found`);
            return;
        }

        const fileInputId = fileInput.id;
        const docConfig = DOCUMENT_TYPE_MAP[fileInputId];

        if (!docConfig) {
            console.error(`Unknown document type for input: ${fileInputId}`);
            return;
        }

        if (!cache.currentId) {
            alert("Please save the DCL first before uploading documents.");
            return;
        }

        const originalValue = textInput.value;
        const fileExtension = getFileExtension(file.name);

        if (!ALLOWED_FILE_EXTENSIONS.includes(fileExtension.toLowerCase())) {
            alert(`Invalid file type. Allowed types: ${ALLOWED_FILE_EXTENSIONS.join(', ').toUpperCase()}`);
            fileInput.value = '';
            return;
        }

        try {
            textInput.value = 'Processing document...';
            textInput.disabled = true;

            // Convert file to base64
            const base64 = await fileToBase64(file);

            // Step 1: Extract number from document
            let extractedNumber = '';
            try {
                textInput.value = 'Extracting number...';

                const extractResponse = await extractNumberFromFile(base64, fileExtension);

                if (extractResponse?.numbers?.length > 0) {
                    if (extractResponse.numbers.length === 1) {
                        const firstItem = extractResponse.numbers[0];
                        extractedNumber = typeof firstItem === 'string'
                            ? firstItem
                            : (firstItem.item || firstItem.number || firstItem.value || '');
                    } else {
                        // Multiple numbers found - create dropdown
                        textInput.value = originalValue;
                        textInput.disabled = false;
                        createNumberDropdownWithStorage(extractResponse.numbers, textInputId, fileInput, base64, fileExtension, docConfig);
                        return;
                    }
                }
            } catch (extractErr) {
                console.warn("Number extraction failed, continuing with upload:", extractErr);
            }

            // Step 2: Upload file to SharePoint
            textInput.value = 'Uploading to SharePoint...';

            const uploadResult = await uploadFileToSharePoint(
                cache.currentId,
                base64,
                docConfig.type,
                fileExtension
            );

            // Get the SharePoint URL from the response
            const sharePointUrl = uploadResult.fileUrl || null;

            // Step 3: Check for existing document record and handle accordingly
            textInput.value = 'Saving document record...';

            const existingDoc = await checkExistingDocument(cache.currentId, docConfig.type);

            if (existingDoc) {
                await deleteDocumentRecord(existingDoc.cr650_dcl_documentid);
            }

            // Create new document record WITH SharePoint URL
            await createDocumentRecord(
                cache.currentId,
                docConfig.type,
                fileExtension,
                `Uploaded: ${file.name}`,
                sharePointUrl  // ← Pass the SharePoint URL
            );

            // Step 4: Update the text field with extracted number
            textInput.value = extractedNumber || '';

            // Step 5: Save extracted number to DCL master if found
            if (extractedNumber && cache.currentId) {
                await patchSingleField(docConfig.field, extractedNumber);
            }

            // Show success feedback
            showUploadSuccess(fileInput, docConfig.type);

            // Log success with URL
            console.log(`✅ Document uploaded successfully. SharePoint URL: ${sharePointUrl || 'N/A'}`);

        } catch (error) {
            console.error('Error processing document:', error);
            alert(`Failed to process document: ${error.message}\n\nPlease try again.`);
            textInput.value = originalValue;
            fileInput.value = '';
        } finally {
            textInput.disabled = false;
        }
    }

    /**
     * Initialize document upload with storage
     */
    function initDocumentNumberExtractionWithStorage() {
        const fileInputMappings = [
            { fileInputId: 'edFile', textInputId: 'edNumber' },
            { fileInputId: 'blFile', textInputId: 'blNumber' },
            { fileInputId: 'piFile', textInputId: 'piNumber' },
            { fileInputId: 'poFile', textInputId: 'poNumber' },
            { fileInputId: 'customerPoFile', textInputId: 'customerPoNumber' },
            { fileInputId: 'lcFile', textInputId: 'lcNumber' }
        ];

        fileInputMappings.forEach(mapping => {
            const fileInput = document.getElementById(mapping.fileInputId);
            if (fileInput) {
                // Remove any existing listeners by cloning
                const newInput = fileInput.cloneNode(true);
                fileInput.parentNode.replaceChild(newInput, fileInput);

                // Add new listener with storage functionality
                newInput.addEventListener('change', function () {
                    handleDocumentFileUploadWithStorage(this, mapping.textInputId);
                });
            }
        });
    }

    async function loadDclById(guid) {
        const url = `${DCL_API}(${encodeURIComponent(guid)})?` +
            "$select=" + encodeURIComponent(selectFields().join(","));
        const data = await safeAjax({ url });
        cache.currentDcl = data;
        cache.original = JSON.parse(JSON.stringify(data || {}));
        return data;
    }

    function setDateInput(selector, iso) {
        const node = $(selector);
        if (!node) return;

        if (!iso) { node.value = ""; return; }

        const dt = new Date(iso);
        if (isNaN(dt)) { node.value = ""; return; }

        const pad = (n) => String(n).padStart(2, "0");
        node.value = [
            dt.getFullYear(),
            pad(dt.getMonth() + 1),
            pad(dt.getDate())
        ].join("-");
    }

    function bindDclToForm(rec) {
        if (!rec) return;

        if ($("#dclNumber")) {
            $("#dclNumber").textContent = rec.cr650_dclnumber || "-";
        }

        const custName = rec.cr650_customername || "";
        const custCode = rec.cr650_customernumber || "";
        trySelectOrInjectCustomer(custName, custCode);

        setVal("#customerNumber", custCode || "");

        const companyType = rec.cr650_company_type;
        if (companyType !== null && companyType !== undefined) {
            setVal("#companyType", String(companyType));
            updateCompanyTypeColor();
        }

        setVal("#country", rec.cr650_country || "");
        setVal("#orgID", rec.cr650_organizationid || "");
        setVal("#COB", rec.cr650_cob || "");
        setVal("#COO", rec.cr650_countryoforigin || "");
        setVal("#PaymentTerms", rec.cr650_paymentterms || "");
        setVal("#currency", rec.cr650_currencycode || "");
        setVal("#incoterms", rec.cr650_incoterms || "");

        setVal("#billTo", rec.cr650_billto || "");
        setVal("#salesRep", rec.cr650_salesrepresentativename || "");
        setVal("#descriptionGoodsServices", rec.cr650_description_goods_services || "");
        setVal("#destinationPort", rec.cr650_destinationport || "");

        setVal("#edNumber", rec.cr650_ednumber || "");
        setVal("#blNumber", rec.cr650_blnumber || "");
        setVal("#piNumber", rec.cr650_pinumber || "");
        setVal("#poNumber", rec.cr650_ponumber || "");
        setVal("#customerPoNumber", rec.cr650_po_customer_number || "");
        setVal("#lcNumber", rec.cr650_lc_number || "");

        setDateInput("#actualRednessDate", rec.cr650_actualreadinessdate);
        setDateInput("#lcDateOfIssue", rec.cr650_lcissuedate);
        setDateInput("#lcLatestShipment", rec.cr650_lclatestshipmentdate);
        setDateInput("#bankSubmissionDate", rec.cr650_banksubmissiondate);
        setDateInput("#lcExpiryDate", rec.cr650_lcexpirydate);
        setDateInput("#loadingDate", rec.cr650_loadingdate);
        setDateInput("#sailingDate", rec.cr650_sailing_date);

        const loadingType = (rec.cr650_loading_type || "").toLowerCase().trim();

        if (loadingType === "shift") {
            $("#loadingTimeType").value = "shift";
            toggleLoadingTimeInput();
            setVal("#loadingShift", rec.cr650_loading_shift || "");
        } else if (loadingType === "specific time" || loadingType === "specific") {
            $("#loadingTimeType").value = "specific";
            toggleLoadingTimeInput();
            setVal("#loadingTime", rec.cr650_loadingtime || "");
        } else {
            if (rec.cr650_loading_shift) {
                $("#loadingTimeType").value = "shift";
                toggleLoadingTimeInput();
                setVal("#loadingShift", rec.cr650_loading_shift || "");
            } else if (rec.cr650_loadingtime) {
                $("#loadingTimeType").value = "specific";
                toggleLoadingTimeInput();
                setVal("#loadingTime", rec.cr650_loadingtime || "");
            } else {
                $("#loadingTimeType").value = "";
                toggleLoadingTimeInput();
            }
        }

        setVal("#shippingLineSelect", rec.cr650_shippingline || "");
        setVal("#loadingPort", rec.cr650_loadingport || "");
        setVal("#transportation", rec.cr650_transportationmode || "");
    }

    const PARTY = {
        LABELS: {
            shipto: "Ship To",
            consignee: "Consignee",
            applicant: "Applicant"
        },
        initialType(rec) {
            const raw = (rec.cr650_party || "").toLowerCase().trim();
            if (raw.includes("consignee")) return "consignee";
            if (raw.includes("applicant")) return "applicant";
            if (raw.includes("ship")) return "shipto";

            if (rec.cr650_consignee) return "consignee";
            if (rec.cr650_shiptolocation) return "shipto";
            if (rec.cr650_applicant) return "applicant";
            return "shipto";
        },
        applyTypeToUI(type, rec) {
            type = (type || "shipto").toLowerCase();
            fillPartyUIFromRecord(type, rec);
        },
        _t: null,
        debouncedPatch(fieldsObj, guid, onOk) {
            clearTimeout(PARTY._t);
            PARTY._t = setTimeout(async () => {
                try {
                    await patchDcl(guid, cleanPatchObject(fieldsObj));
                    Object.assign(cache.original, fieldsObj);
                    Object.assign(cache.currentDcl, fieldsObj);
                    if (typeof onOk === "function") onOk();
                } catch (e) {
                    console.error("Party PATCH failed", e);
                    alert("Failed to save party change. Check permissions or connection.");
                }
            }, 500);
        }
    };

    function collectPayloadFromForm() {
        const partyType = ($("#shipToLabel")?.value || "shipto").toLowerCase();
        const mainText = $("#shipTo")?.value || "";
        const billTo = $("#billTo")?.value || "";
        const loadingTypeSel = ($("#loadingTimeType")?.value || "").toLowerCase();

        const base = {
            cr650_customername: cache.selectedCustomer?.name || cache.currentDcl?.cr650_customername || "",
            cr650_customernumber: $("#customerNumber")?.value || "",
            cr650_company_type: $("#companyType")?.value ? parseInt($("#companyType").value) : null,
            cr650_country: $("#country")?.value || "",
            cr650_organizationid: $("#orgID")?.value || "",
            cr650_cob: $("#COB")?.value || "",
            cr650_countryoforigin: $("#COO")?.value || "",
            cr650_paymentterms: $("#PaymentTerms")?.value || "",
            cr650_currencycode: $("#currency")?.value || "",
            cr650_incoterms: $("#incoterms")?.value || "",

            cr650_ednumber: $("#edNumber")?.value || "",
            cr650_blnumber: $("#blNumber")?.value || "",
            cr650_pinumber: $("#piNumber")?.value || "",
            cr650_ponumber: $("#poNumber")?.value || "",
            cr650_po_customer_number: $("#customerPoNumber")?.value || "",
            cr650_lc_number: $("#lcNumber")?.value || "",

            cr650_actualreadinessdate: asISO($("#actualRednessDate")?.value),
            cr650_lcissuedate: asISO($("#lcDateOfIssue")?.value),
            cr650_lclatestshipmentdate: asISO($("#lcLatestShipment")?.value),
            cr650_banksubmissiondate: asISO($("#bankSubmissionDate")?.value),
            cr650_lcexpirydate: asISO($("#lcExpiryDate")?.value),

            cr650_shippingline: $("#shippingLineSelect")?.value || "",
            cr650_loadingport: $("#loadingPort")?.value || "",
            cr650_destinationport: $("#destinationPort")?.value || "",
            cr650_loadingdate: asISO($("#loadingDate")?.value),

            cr650_loadingtime: $("#loadingTime")?.value || "",
            cr650_loading_shift: $("#loadingShift")?.value || "",

            cr650_loading_type:
                loadingTypeSel === "shift"
                    ? "Shift"
                    : loadingTypeSel === "specific"
                        ? "Specific Time"
                        : cache.currentDcl?.cr650_loading_type || "",

            cr650_transportationmode: $("#transportation")?.value || "",
            cr650_salesrepresentativename: $("#salesRep")?.value || "",
            cr650_party: PARTY.LABELS[partyType] || "Ship To"
        };

        if (partyType === "consignee") {
            base.cr650_consignee = mainText;
        } else if (partyType === "applicant") {
            base.cr650_applicant = mainText;
        } else {
            base.cr650_shiptolocation = mainText;
            base.cr650_billto = billTo;
        }

        Object.keys(base).forEach(k => {
            if (base[k] === "" || base[k] == null) delete base[k];
        });

        return base;
    }

    function diffPayload(newObj, oldObj) {
        const out = {};
        Object.keys(newObj).forEach(k => {
            const nv = newObj[k];
            const ov = oldObj[k];
            if (nv !== undefined && nv !== ov) out[k] = nv;
        });
        return out;
    }

    async function patchDcl(guid, bodyObj) {
        if (!guid) throw new Error("Missing DCL id");

        const url = `${DCL_API}(${encodeURIComponent(guid)})`;

        return safeAjax({
            type: "PATCH",
            url,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                "If-Match": "*"
            },
            data: JSON.stringify(bodyObj)
        });
    }

    async function patchSingleField(attrName, rawVal) {
        if (!cache.currentId) return;
        const body = {};
        if (rawVal !== "" && rawVal !== undefined && rawVal !== null) {
            body[attrName] = rawVal;
        } else {
            body[attrName] = null;
        }

        try {
            await patchDcl(cache.currentId, body);
            cache.original[attrName] = body[attrName];
            cache.currentDcl[attrName] = body[attrName];
        } catch (err) {
            console.error("PATCH failed for", attrName, err);
            alert("Failed to save changes to server.");
        }
    }

    async function onSaveClick() {
        if (!cache.currentId) {
            alert("Missing DCL id in URL.");
            return;
        }

        const proposed = collectPayloadFromForm();
        const delta = diffPayload(proposed, cache.original || {});
        if (!Object.keys(delta).length) {
            w.location.href = "/loading_plan_view/?id=" + encodeURIComponent(cache.currentId);
            return;
        }

        setBusy(true, "Saving…");
        try {
            await patchDcl(cache.currentId, delta);
            Object.assign(cache.original, delta);
            Object.assign(cache.currentDcl, delta);
            w.location.href = "/loading_plan_view/?id=" + encodeURIComponent(cache.currentId);
        } catch (err) {
            console.error("Failed to update DCL", err);
            alert("Error updating DCL. Check required fields and permissions.");
        } finally {
            setBusy(false);
        }
    }

    /* ============================================================
       12) COLLAPSIBLES / UI CONTROLS
       ============================================================ */
    function initCollapsibles() {
        $$(".info-section").forEach(sec => {
            const btn = sec.querySelector(".section-header");
            const content = sec.querySelector(".section-content");
            const isOpen = (btn?.getAttribute("aria-expanded") || "true") === "true";

            sec.classList.toggle("open", isOpen);
            if (content) {
                content.style.display = isOpen ? "block" : "none";
            }

            const ic = sec.querySelector(".toggle-icon");
            if (ic) ic.style.transform = isOpen ? "rotate(90deg)" : "rotate(0deg)";
        });

        d.addEventListener("click", (e) => {
            const header = e.target.closest(".section-header");
            if (!header) return;
            const card = header.closest(".info-section");
            const content = card?.querySelector(".section-content");
            if (!content || !content.id) return;
            w.wizard.toggleInfoSection(content.id);
        });
    }

    w.wizard = w.wizard || {};
    w.wizard.toggleInfoSection = function (sectionId) {
        const content = d.getElementById(sectionId);
        if (!content) return;

        const card = content.closest(".info-section");
        const icon = d.getElementById(`${sectionId}-icon`);
        const btn = card ? card.querySelector(".section-header") : null;

        const willOpen = !(card && card.classList.contains("open"));
        if (card) card.classList.toggle("open", willOpen);

        content.style.display = willOpen ? "block" : "none";
        if (btn) btn.setAttribute("aria-expanded", String(willOpen));
        if (icon) icon.style.transform = willOpen ? "rotate(90deg)" : "rotate(0deg)";
    };

    w.handleShippingLineChange = async function (selectEl) {
        if (!selectEl) return;

        if (selectEl.value === "__add_new__") {
            const name = prompt("Enter new shipping line name:");

            if (name && name.trim()) {
                const clean = name.trim();

                try {
                    TopLoader.begin();
                    const newLine = await createShippingLine(clean);

                    const allLines = await loadAllShippingLines();
                    populateShippingLineSelect(selectEl, allLines);

                    selectEl.value = newLine.name;

                    if (cache.currentId) {
                        await patchSingleField("cr650_shippingline", newLine.name);
                    }

                    TopLoader.end();
                } catch (err) {
                    TopLoader.end();
                    console.error("Failed to add shipping line", err);
                    alert("Failed to add new shipping line. Please try again.");
                    selectEl.value = "";
                }
            } else {
                selectEl.value = "";
            }
        }
    };

    function toggleLoadingTimeInput() {
        const type = $("#loadingTimeType")?.value || "";
        const specific = $("#specificTimeInput");
        const shift = $("#shiftInput");

        if (specific) specific.style.display = (type === "specific") ? "" : "none";
        if (shift) shift.style.display = (type === "shift") ? "" : "none";
    }
    w.toggleLoadingTimeInput = toggleLoadingTimeInput;

    w.handleBrandChange = async function (selectEl) {
        if (!selectEl) return;
        const listEl = $("#brandList");

        if (selectEl.value === "__add_new__") {
            const name = prompt("Enter new brand name:");

            if (name && name.trim()) {
                const clean = name.trim();

                try {
                    TopLoader.begin();

                    // Create the brand in the lookup list
                    await createBrandInList(clean);

                    // Reload and repopulate the brand select
                    const allBrands = await loadAllBrandsList();
                    cache.brandsList = allBrands;
                    populateBrandSelect(selectEl, allBrands);

                    // Set the newly added brand as selected
                    selectEl.value = clean;

                    // Also add it to the DCL brands
                    await createBrand(clean);
                    const refreshedBrands = await loadBrands(cache.currentId);
                    renderBrandList(refreshedBrands);

                    TopLoader.end();
                } catch (err) {
                    TopLoader.end();
                    console.error("Failed to add new brand", err);
                    alert("Failed to add new brand. Please try again.");
                    selectEl.value = "";
                }
            } else {
                selectEl.value = "";
            }
            return;
        }

        // Handle selecting an existing brand — auto-add it (no separate Add button needed)
        const value = selectEl.value;
        const text = selectEl.options[selectEl.selectedIndex]?.textContent?.trim() || "";
        if (!value || !text) return;

        try {
            await createBrand(text);
            const refreshedBrands = await loadBrands(cache.currentId);
            renderBrandList(refreshedBrands);
        } catch (errBrand) {
            console.error("Failed to add brand", errBrand);
            alert("Failed to add Brand. Check permissions or connection.");
        }

        selectEl.value = "";
    };

    // Keep backward compat alias
    w.onAddBrand = function () {
        const select = $("#brandSelect");
        if (select) w.handleBrandChange(select);
    };

    w.onAddOrderNumber = async function () {
        const sel = $("#orderNumberSelect");
        const tableSection = $("#orderNumbersTableSection");
        const tbody = $("#orderNumbersTable tbody");
        if (!sel || !tbody || !tableSection) return;

        const order = sel.value;
        const selectedOption = sel.options[sel.selectedIndex];
        const label = selectedOption?.textContent || "";
        const custPoNumber = selectedOption?.dataset.custPoNumber || "";
        if (!order) return;

        if ([...tbody.querySelectorAll("tr")].some(tr => tr.dataset.order === order)) return;

        try {
            const createdRow = await createDclOrder(order, custPoNumber);
            const dclOrderId = createdRow.cr650_dcl_orderid || "";

            const tr = d.createElement("tr");
            tr.dataset.order = order;
            tr.dataset.dclOrderId = dclOrderId;

            const tdOrder = d.createElement("td");
            tdOrder.style.padding = "10px";
            tdOrder.textContent = label;

            const tdCI = d.createElement("td");
            tdCI.style.padding = "10px";
            tdCI.textContent = custPoNumber || "";

            const tdAct = d.createElement("td");
            tdAct.style.padding = "10px";
            const rm = d.createElement("button");
            rm.type = "button";
            rm.textContent = "Remove";
            rm.className = "btn btn-outline";
            rm.addEventListener("click", async () => {
                const sure = confirm("Remove this order from this DCL?");
                if (!sure) return;

                try {
                    if (dclOrderId) {
                        await deleteDclOrder(dclOrderId);
                    }
                    tr.remove();
                    if (!tbody.children.length) {
                        tableSection.style.display = "none";
                    }
                } catch (errDelOne) {
                    console.error("Failed to delete dcl_order row", errDelOne);
                    alert("Couldn't remove this order from the DCL.");
                }
            });
            tdAct.appendChild(rm);

            tr.appendChild(tdOrder);
            tr.appendChild(tdCI);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);

            tableSection.style.display = "";
            sel.value = "";
        } catch (errCreate) {
            console.error("Failed to create DCL order row", errCreate);
            alert("Failed to add order to DCL. Check permissions or connection.");
        }
    };

    /* ============================================================
       13) GENERIC SETVAL
       ============================================================ */
    function setVal(sel, v) {
        const node = $(sel);
        if (!node) return;
        const tag = node.tagName.toLowerCase();
        if (["input", "select", "textarea"].includes(tag)) {
            node.value = v ?? "";
        } else {
            node.textContent = v ?? "";
        }
    }

    /* ============================================================
       14) CUSTOMER CHANGE HANDLER (from picker)
       ============================================================ */
    function computePartyFieldsFromCustomer(cust) {
        let partyType = "shipto";
        if (cust.shipTo) {
            partyType = "shipto";
        } else if (cust.consignee) {
            partyType = "consignee";
        } else if (cust.primaryAddress) {
            partyType = "applicant";
        }

        const patch = {};
        if (partyType === "shipto") {
            patch.cr650_party = "Ship To";
            patch.cr650_shiptolocation = cust.shipTo || "";
            patch.cr650_billto = cust.billTo || "";
            patch.cr650_consignee = "";
            patch.cr650_applicant = "";
        } else if (partyType === "consignee") {
            patch.cr650_party = "Consignee";
            patch.cr650_consignee = cust.consignee || "";
            patch.cr650_shiptolocation = "";
            patch.cr650_billto = "";
            patch.cr650_applicant = "";
        } else {
            patch.cr650_party = "Applicant";
            patch.cr650_applicant = cust.primaryAddress || "";
            patch.cr650_shiptolocation = "";
            patch.cr650_billto = "";
            patch.cr650_consignee = "";
        }

        return { patch, partyType };
    }

    function applyShipToUIFromPartyType(partyType, custRecord) {
        fillPartyUIFromRecord(partyType, custRecord);
    }

    // Lock to prevent multiple concurrent handleCustomerChange calls
    let customerChangeLock = false;

    async function handleCustomerChange() {
        // Prevent concurrent calls
        if (customerChangeLock) {
            console.log("⏳ Customer change already in progress, skipping...");
            return;
        }
        customerChangeLock = true;

        try {
            const selNode = el.customerPicker;
            const idAttr = selNode?.value || "";

            if (!idAttr) {
                cache.selectedCustomer = null;
                cache.customerAvailableModels = [];

                setVal("#customerNameText", "");
                setVal("#customerNumber", "");
                setVal("#country", "");
                setVal("#orgID", "");
                setVal("#COB", "");
                setVal("#PaymentTerms", "");
                setVal("#salesRep", "");
                setVal("#destinationPort", "");
                setVal("#shipTo", "");
                setVal("#billTo", "");
                clearOrderSelectAndTable();
                renderCustomerAvailableModels([]);
                customerChangeLock = false;
                return;
            }

        const idx = selNode.options[selNode.selectedIndex]?.dataset.idx;
        const cust =
            cache.customers[idx] ||
            cache.customers.find(c => c.id === idAttr) ||
            null;

        cache.selectedCustomer = cust;
        if (!cust) {
            clearOrderSelectAndTable();
            customerChangeLock = false;
            return;
        }

        if (el.customerNameText) {
            el.customerNameText.value = cust.code
                ? `${cust.name} (${cust.code})`
                : (cust.name || "");
        }

        setVal("#customerNumber", cust.code || "");
        setVal("#country", cust.country || "");
        setVal("#orgID", cust.orgId || "");
        setVal("#COB", cust.cob || "");
        setVal("#PaymentTerms", cust.paymentTerms || "");
        setVal("#salesRep", cust.salesRep || "");
        setVal("#destinationPort", cust.destinationPort || "");

        const { patch: partyPatch, partyType } = computePartyFieldsFromCustomer(cust);
        applyShipToUIFromPartyType(partyType, cust);
        tryDefaultCurrencyByCountry();

        // Load and display customer's available models
        try {
            const availableModels = await loadCustomerAvailableModels(cust.id);
            cache.customerAvailableModels = availableModels;
            renderCustomerAvailableModels(availableModels);
        } catch (err) {
            console.error("Failed to load customer available models:", err);
            cache.customerAvailableModels = [];
            renderCustomerAvailableModels([]);
        }

        const custNumber = (cust.code || "").trim();

        if (!cache.currentId) {
            console.warn("No DCL ID in cache.currentId, skip PATCH.");
            // Still reload orders for the new customer even without DCL ID
            if (custNumber) {
                await reloadOrdersForCustomerNumber(custNumber);
            } else {
                clearOrderSelectAndTable();
            }
            customerChangeLock = false;
            return;
        }

        const patchBodyRaw = {
            cr650_customername: cust.name || "",
            cr650_customernumber: cust.code || "",
            cr650_country: cust.country || "",
            cr650_organizationid: cust.orgId || "",
            cr650_cob: cust.cob || "",
            cr650_paymentterms: cust.paymentTerms || "",
            cr650_salesrepresentativename: cust.salesRep || "",
            cr650_destinationport: cust.destinationPort || "",
            ...partyPatch
        };

        const cleanedPatch = cleanPatchObject(patchBodyRaw);

        console.log("📤 Customer change PATCH payload:", cleanedPatch);

        try {
            console.log("📤 Customer change PATCH JSON:", JSON.stringify(cleanedPatch));
            await patchDcl(cache.currentId, cleanedPatch);
            Object.assign(cache.original, cleanedPatch);
            Object.assign(cache.currentDcl, cleanedPatch);
        } catch (errPatch) {
            console.error("Failed to PATCH customer switch - Full error:", errPatch);
            console.error("Error details:", {
                code: errPatch?.error?.code,
                message: errPatch?.error?.message,
                innererror: errPatch?.error?.innererror,
                fullError: JSON.stringify(errPatch, null, 2)
            });
            const errorMsg = errPatch?.error?.message || errPatch?.message || "Unknown error";
            const errorCode = errPatch?.error?.code ? ` (Code: ${errPatch.error.code})` : "";
            alert("Couldn't update the DCL with the selected customer.\n\nError: " + errorMsg + errorCode);
            customerChangeLock = false;
            return;
        }

        try {
            await deleteAllOrdersForCurrentDcl();
        } catch (bulkErr) {
            console.error("Failed bulk delete of orders after customer change", bulkErr);
        }

        // Clear the orders table (old DCL orders were deleted)
        clearOrderTableOnly();

        // Reload available orders for the new customer
        if (custNumber) {
            await reloadOrdersForCustomerNumber(custNumber);
        } else {
            clearOrderSelectAndTable();
        }

        // Auto-populate notify parties from customer data (add as drafts, keep existing)
        try {
            const np1 = (cust.notifyParty1 || "").trim();
            const np2 = (cust.notifyParty2 || "").trim();

            // Only add if customer has default notify parties
            if (np1 || np2) {
                const container = d.getElementById("notifyPartyContainer");
                if (container) {
                    // Add notify party 1 as a draft row (user needs to click Save)
                    if (np1) {
                        const draftItem1 = {
                            id: "",
                            text: np1,
                            created: "",
                            dclGuid: cache.currentId || "",
                            dclNumber: ""
                        };
                        const rowEl1 = buildNotifyPartyRow(draftItem1, true);
                        container.appendChild(rowEl1);
                    }

                    // Add notify party 2 as a draft row (user needs to click Save)
                    if (np2) {
                        const draftItem2 = {
                            id: "",
                            text: np2,
                            created: "",
                            dclGuid: cache.currentId || "",
                            dclNumber: ""
                        };
                        const rowEl2 = buildNotifyPartyRow(draftItem2, true);
                        container.appendChild(rowEl2);
                    }
                }
            }
        } catch (npErr) {
            console.error("Failed to add notify party drafts from customer", npErr);
        }

        if (el.customerPicker) {
            el.customerPicker.classList.add("hidden");
        }

        if (el.btnPickCustomer) {
            el.btnPickCustomer.textContent = "Change";
        }
        } finally {
            // Always release the lock
            customerChangeLock = false;
        }
    }

    function onShipToLabelChange() {
        const mode = ($("#shipToLabel")?.value || "").toLowerCase();

        const sourceRecord = (cache.selectedCustomer && cache.selectedCustomer.id)
            ? cache.selectedCustomer
            : cache.currentDcl;

        fillPartyUIFromRecord(mode, sourceRecord);

        const mainText = $("#shipTo")?.value || "";
        const billText = $("#billTo")?.value || "";

        let fields = {};
        if (mode === "consignee") {
            fields = {
                cr650_party: "Consignee",
                cr650_consignee: mainText,
                cr650_shiptolocation: "",
                cr650_billto: "",
                cr650_applicant: ""
            };
        } else if (mode === "applicant") {
            fields = {
                cr650_party: "Applicant",
                cr650_applicant: mainText,
                cr650_shiptolocation: "",
                cr650_consignee: "",
                cr650_billto: ""
            };
        } else {
            fields = {
                cr650_party: "Ship To",
                cr650_shiptolocation: mainText,
                cr650_billto: billText,
                cr650_consignee: "",
                cr650_applicant: ""
            };
        }

        PARTY.debouncedPatch(fields, cache.currentId, () => {
            Object.assign(cache.original, fields);
            Object.assign(cache.currentDcl, fields);
        });
    }

    /* ============================================================
       15) GLOBAL CLICK HANDLERS (Notify/T&C/Brand remove)
       ============================================================ */
    d.addEventListener("click", async (e) => {

        // Notify Save
        if (e.target.classList.contains("notify-save-btn")) {
            const card = e.target.closest(".notify-card");
            if (!card) return;

            const textarea = card.querySelector(".notify-text");
            const noteText = (textarea?.value || "").trim();
            const id = card.dataset.notifyId || "";
            const mode = e.target.getAttribute("data-mode") || (id ? "update" : "create");

            if (!noteText) {
                if (!id) card.remove();
                return;
            }

            // Collect other unsaved draft cards before saving (to preserve them)
            const container = d.getElementById("notifyPartyContainer");
            const otherDrafts = [];
            if (container) {
                container.querySelectorAll('.notify-card[data-notify-id=""]').forEach(draftCard => {
                    if (draftCard !== card) {
                        const draftText = (draftCard.querySelector(".notify-text")?.value || "").trim();
                        if (draftText) {
                            otherDrafts.push(draftText);
                        }
                    }
                });
            }

            try {
                if (mode === "create" || !id) {
                    await createNotifyParty(noteText);
                } else {
                    await updateNotifyParty(id, noteText);
                }

                const refreshed = await loadNotifyParties(cache.currentId);
                renderNotifyPartyList(refreshed);

                // Re-add other draft cards that weren't saved yet
                if (container && otherDrafts.length > 0) {
                    otherDrafts.forEach(draftText => {
                        const draftItem = {
                            id: "",
                            text: draftText,
                            created: "",
                            dclGuid: cache.currentId || "",
                            dclNumber: ""
                        };
                        const rowEl = buildNotifyPartyRow(draftItem, true);
                        container.appendChild(rowEl);
                    });
                }
            } catch (errSave) {
                console.error("Failed to save notify party", errSave);
                alert("Failed to save Notify Party. Check permissions or connection.");
            }
            return;
        }

        // Notify Delete
        if (e.target.classList.contains("notify-delete-btn")) {
            const card = e.target.closest(".notify-card");
            if (!card) return;

            const id = card.dataset.notifyId || "";

            if (!id) {
                card.remove();
                return;
            }

            const sure = confirm("Delete this notify party?");
            if (!sure) return;

            try {
                await deleteNotifyParty(id);
                const refreshed = await loadNotifyParties(cache.currentId);
                renderNotifyPartyList(refreshed);
            } catch (errDel) {
                console.error("Failed to delete notify party", errDel);
                alert("Failed to delete Notify Party. Check permissions or connection.");
            }
            return;
        }

        // Customer Model Save
        if (e.target.classList.contains("model-save-btn") || e.target.closest(".model-save-btn")) {
            const btn = e.target.classList.contains("model-save-btn") ? e.target : e.target.closest(".model-save-btn");
            const card = btn.closest(".customer-model-card");
            if (!card) return;

            const nameInput = card.querySelector(".model-name-input");
            const infoInput = card.querySelector(".model-info-input");
            const modelName = (nameInput?.value || "").trim();
            const modelInfo = (infoInput?.value || "").trim();
            const id = card.dataset.modelId || "";
            const mode = btn.getAttribute("data-mode") || (id ? "update" : "create");

            if (!modelName) {
                alert("Model name is required.");
                if (nameInput) nameInput.focus();
                return;
            }

            // Collect other unsaved draft cards before saving (to preserve them)
            const container = d.getElementById("customerModelContainer");
            const otherDrafts = [];
            if (container) {
                container.querySelectorAll('.customer-model-card[data-model-id=""]').forEach(draftCard => {
                    if (draftCard !== card) {
                        const draftName = (draftCard.querySelector(".model-name-input")?.value || "").trim();
                        const draftInfo = (draftCard.querySelector(".model-info-input")?.value || "").trim();
                        if (draftName) {
                            otherDrafts.push({ name: draftName, info: draftInfo });
                        }
                    }
                });
            }

            try {
                if (mode === "create" || !id) {
                    await createCustomerModel(modelName, modelInfo);
                } else {
                    await updateCustomerModel(id, modelName, modelInfo);
                }

                const refreshed = await loadCustomerModels(cache.currentId);
                renderCustomerModelList(refreshed);

                // Re-add other draft cards that weren't saved yet
                if (container && otherDrafts.length > 0) {
                    otherDrafts.forEach(draft => {
                        appendDraftCustomerModelRow(draft.name, draft.info);
                    });
                }
            } catch (errSave) {
                console.error("Failed to save customer model", errSave);
                alert("Failed to save Customer Model. Check permissions or connection.");
            }
            return;
        }

        // Customer Model Delete
        if (e.target.classList.contains("model-delete-btn") || e.target.closest(".model-delete-btn")) {
            const btn = e.target.classList.contains("model-delete-btn") ? e.target : e.target.closest(".model-delete-btn");
            const card = btn.closest(".customer-model-card");
            if (!card) return;

            const id = card.dataset.modelId || "";

            if (!id) {
                card.remove();
                return;
            }

            const sure = confirm("Delete this customer model?");
            if (!sure) return;

            try {
                await deleteCustomerModel(id);
                const refreshed = await loadCustomerModels(cache.currentId);
                renderCustomerModelList(refreshed);
            } catch (errDel) {
                console.error("Failed to delete customer model", errDel);
                alert("Failed to delete Customer Model. Check permissions or connection.");
            }
            return;
        }

        // Brand chip remove
        if (e.target.classList.contains("brand-remove-btn")) {
            const chip = e.target.closest(".brand-chip");
            if (!chip) return;

            const brandId = chip.dataset.brandId || "";

            if (!brandId) {
                chip.remove();
                return;
            }

            const sure = confirm("Remove this brand from this DCL?");
            if (!sure) return;

            try {
                await deleteBrand(brandId);
                const refreshedBrands = await loadBrands(cache.currentId);
                renderBrandList(refreshedBrands);
            } catch (errDelBrand) {
                console.error("Failed to delete brand", errDelBrand);
                alert("Failed to delete Brand. Check permissions or connection.");
            }

            return;
        }
    });

    /* ============================================================
       16) AUTOSAVE FIELD BINDINGS
       ============================================================ */
    const FIELD_BINDINGS = [
        { selector: "#country", attr: "cr650_country", event: "blur" },
        {
            selector: "#companyType",
            attr: "cr650_company_type",
            event: "change",
            transform: (val) => val ? parseInt(val) : null,
            afterSave: () => {
                updateCompanyTypeColor();
            }
        },
        { selector: "#orgID", attr: "cr650_organizationid", event: "blur" },
        { selector: "#COB", attr: "cr650_cob", event: "blur" },
        { selector: "#COO", attr: "cr650_countryoforigin", event: "blur" },
        { selector: "#PaymentTerms", attr: "cr650_paymentterms", event: "blur" },
        { selector: "#currency", attr: "cr650_currencycode", event: "change" },
        { selector: "#incoterms", attr: "cr650_incoterms", event: "change" },

        { selector: "#edNumber", attr: "cr650_ednumber", event: "blur" },
        { selector: "#blNumber", attr: "cr650_blnumber", event: "blur" },
        { selector: "#piNumber", attr: "cr650_pinumber", event: "blur" },
        { selector: "#poNumber", attr: "cr650_ponumber", event: "blur" },
        { selector: "#customerPoNumber", attr: "cr650_po_customer_number", event: "blur" },
        { selector: "#lcNumber", attr: "cr650_lc_number", event: "blur" },

        { selector: "#actualRednessDate", attr: "cr650_actualreadinessdate", event: "change", transform: asISO },
        { selector: "#lcDateOfIssue", attr: "cr650_lcissuedate", event: "change", transform: asISO },
        { selector: "#lcLatestShipment", attr: "cr650_lclatestshipmentdate", event: "change", transform: asISO },
        { selector: "#bankSubmissionDate", attr: "cr650_banksubmissiondate", event: "change", transform: asISO },
        { selector: "#lcExpiryDate", attr: "cr650_lcexpirydate", event: "change", transform: asISO },

        { selector: "#shippingLineSelect", attr: "cr650_shippingline", event: "change" },
        { selector: "#loadingPort", attr: "cr650_loadingport", event: "blur" },
        { selector: "#destinationPort", attr: "cr650_destinationport", event: "blur" },

        { selector: "#loadingDate", attr: "cr650_loadingdate", event: "change", transform: asISO },
        { selector: "#sailingDate", attr: "cr650_sailing_date", event: "change", transform: asISO },

        {
            selector: "#loadingTimeType",
            attr: "cr650_loading_type",
            event: "change",
            transform: (val) => {
                if (val === "shift") return "Shift";
                if (val === "specific") return "Specific Time";
                return null;
            },
            afterSave: async (rawVal) => {
                const val = (rawVal || "").toLowerCase();

                if (val === "shift") {
                    await patchSingleField("cr650_loadingtime", null);
                } else if (val === "specific") {
                    await patchSingleField("cr650_loading_shift", null);
                } else {
                    await patchSingleField("cr650_loadingtime", null);
                    await patchSingleField("cr650_loading_shift", null);
                }
            }
        },

        { selector: "#loadingTime", attr: "cr650_loadingtime", event: "change" },
        { selector: "#loadingShift", attr: "cr650_loading_shift", event: "change" },

        { selector: "#transportation", attr: "cr650_transportationmode", event: "change" },
        { selector: "#salesRep", attr: "cr650_salesrepresentativename", event: "blur" },
        { selector: "#descriptionGoodsServices", attr: "cr650_description_goods_services", event: "blur" },

        {
            selector: "#customerNumber",
            attr: "cr650_customernumber",
            event: "blur",
            afterSave: async (valRaw) => {
                const newCustNum = (valRaw || "").trim();

                try {
                    await deleteAllOrdersForCurrentDcl();
                } catch (bulkErr) {
                    console.error("Failed bulk delete after manual cust number change", bulkErr);
                }

                clearOrderTableOnly();

                await reloadOrdersForCustomerNumber(newCustNum);
            }
        },
    ];

    function wireAutoSaveFields() {
        FIELD_BINDINGS.forEach(binding => {
            const fieldEl = $(binding.selector);
            if (!fieldEl) return;

            const handler = async () => {
                const rawVal = fieldEl.value;
                const val = binding.transform ? binding.transform(rawVal) : rawVal;

                await patchSingleField(binding.attr, val);

                if (binding.afterSave) {
                    try {
                        await binding.afterSave(rawVal);
                    } catch (e) {
                        console.error("afterSave error for", binding.selector, e);
                    }
                }
            };

            fieldEl.addEventListener(binding.event, handler);
        });

        const shipToLabelSel = document.getElementById("shipToLabel");
        if (shipToLabelSel) {
            shipToLabelSel.addEventListener("change", onShipToLabelChange);
        }

        const mainTA = document.getElementById("shipTo");
        if (mainTA) {
            mainTA.addEventListener("input", (e) => {
                const chosen =
                    (document.getElementById("shipToLabel")?.value || "")
                        .toLowerCase();
                const val = e.target.value || "";

                let fields = {};
                if (chosen === "consignee") {
                    fields = {
                        cr650_consignee: val,
                        cr650_shiptolocation: "",
                        cr650_applicant: "",
                        cr650_billto: "",
                        cr650_party: PARTY.LABELS[chosen] || "Consignee"
                    };
                } else if (chosen === "applicant") {
                    fields = {
                        cr650_applicant: val,
                        cr650_shiptolocation: "",
                        cr650_consignee: "",
                        cr650_billto: "",
                        cr650_party: PARTY.LABELS[chosen] || "Applicant"
                    };
                } else {
                    fields = {
                        cr650_shiptolocation: val,
                        cr650_consignee: "",
                        cr650_applicant: "",
                        cr650_party: PARTY.LABELS[chosen] || "Ship To"
                    };
                }

                PARTY.debouncedPatch(fields, cache.currentId, () => {
                    Object.assign(cache.original, fields);
                    Object.assign(cache.currentDcl, fields);
                });
            });
        }

        const billToTA = document.getElementById("billTo");
        if (billToTA) {
            billToTA.addEventListener("input", (e) => {
                const chosen =
                    (document.getElementById("shipToLabel")?.value || "")
                        .toLowerCase();
                if (chosen !== "shipto") return;

                const fields = { cr650_billto: e.target.value || "" };
                PARTY.debouncedPatch(fields, cache.currentId, () => {
                    Object.assign(cache.original, fields);
                    Object.assign(cache.currentDcl, fields);
                });
            });
        }
    }

    /* ============================================================
       18A) LOCK FORM IF SUBMITTED
       ============================================================ */
    function lockFormIfSubmitted() {
        try {
            const status = (cache.currentDcl?.cr650_status || "").toLowerCase();

            if (status !== "submitted") {
                console.log("📝 Form is editable - status:", cache.currentDcl?.cr650_status || "none");
                return;
            }

            console.log("🔒 Locking form - DCL status is 'Submitted'");

            const editableFields = $$("input:not([type='radio']):not([type='checkbox']), textarea, select");
            editableFields.forEach(field => {
                field.disabled = true;
                field.readOnly = true;
                field.style.cursor = "not-allowed";
                field.style.opacity = "0.6";
                field.style.backgroundColor = "#f5f5f5";
                field.style.pointerEvents = "none";
            });

            const buttons = $$("button");
            buttons.forEach(btn => {
                if (!btn.closest(".step-indicators")) {
                    btn.disabled = true;
                    btn.style.cursor = "not-allowed";
                    btn.style.opacity = "0.5";
                    btn.style.pointerEvents = "none";
                }
            });

            const actionButtons = [
                "#addOrderBtn",
                "#btnAddNotifyParty",
                "#addNotifyPartyBtn"
            ];

            actionButtons.forEach(selector => {
                const btn = $(selector);
                if (btn) btn.style.display = "none";
            });

            $$(".brand-remove-btn, .notify-delete-btn, .notify-save-btn").forEach(btn => {
                btn.style.display = "none";
            });

            $$(".notify-buttons").forEach(div => {
                div.style.display = "none";
            });

            if (el.btnPickCustomer) {
                el.btnPickCustomer.disabled = true;
                el.btnPickCustomer.style.cursor = "not-allowed";
                el.btnPickCustomer.style.opacity = "0.5";
                el.btnPickCustomer.style.pointerEvents = "none";
            }

            if (el.customerPicker) {
                el.customerPicker.disabled = true;
                el.customerPicker.style.display = "none";
            }

            const brandSelect = $("#brandSelect");
            if (brandSelect) {
                brandSelect.disabled = true;
                brandSelect.style.pointerEvents = "none";
            }

            const orderSelect = $("#orderNumberSelect");
            if (orderSelect) {
                orderSelect.disabled = true;
                orderSelect.style.pointerEvents = "none";
            }

            $$(".brand-chip").forEach(chip => {
                chip.style.cursor = "not-allowed";
                chip.style.opacity = "0.7";
            });

            $$(".notify-card").forEach(card => {
                card.style.opacity = "0.7";
                card.style.pointerEvents = "none";
            });

            showLockedBanner();

            $$("input[type='file']").forEach(fileInput => {
                fileInput.disabled = true;
                fileInput.style.cursor = "not-allowed";
                fileInput.style.opacity = "0.5";
            });

            console.log("✅ Form fully locked");

        } catch (lockError) {
            console.error("❌ Error while locking form:", lockError);
        }
    }

    function showLockedBanner() {
        try {
            const wizard = $("#dclWizard");
            if (!wizard) {
                console.warn("Wizard container not found");
                return;
            }

            if ($("#lockedBanner")) {
                console.log("Banner already exists");
                return;
            }

            const banner = d.createElement("div");
            banner.id = "lockedBanner";
            banner.style.cssText = `
                background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%);
                color: white;
                padding: 16px 24px;
                border-radius: 12px;
                margin-bottom: 20px;
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
                        This DCL has been submitted and cannot be edited.
                    </div>
                </div>
            `;

            const firstSection = wizard.querySelector(".step-content");
            if (firstSection) {
                wizard.insertBefore(banner, firstSection);
            } else {
                wizard.insertBefore(banner, wizard.firstChild);
            }
        } catch (bannerError) {
            console.error("Error creating locked banner:", bannerError);
        }
    }

    /* ============================================================
       17) SAVE BUTTON (Next)
       ============================================================ */
    w.saveDclMaster = onSaveClick;

    /* ============================================================
       15.5) COMPANY TYPE COLOR HANDLER
       ============================================================ */
    function updateCompanyTypeColor() {
        const select = $("#companyType");
        const icon = $("#companyTypeIcon");

        if (!select) return;

        select.classList.remove("petrolube-selected", "technolube-selected");

        if (icon) {
            icon.style.display = "none";
        }

        const value = select.value;

        if (value === "1") {
            select.classList.add("petrolube-selected");
            if (icon) {
                icon.textContent = "✓";
                icon.style.color = "#16a34a";
                icon.style.display = "block";
            }
        } else if (value === "0") {
            select.classList.add("technolube-selected");
            if (icon) {
                icon.textContent = "✓";
                icon.style.color = "#dc2626";
                icon.style.display = "block";
            }
        }
    }

    w.updateCompanyTypeColor = updateCompanyTypeColor;

    /* ============================================================
       18) BOOTSTRAP
       ============================================================ */
    d.addEventListener("DOMContentLoaded", async () => {

        el.customerPicker = $("#customerPicker");
        el.customerNameText = $("#customerNameText");
        el.companyType = $("#companyType");
        el.btnPickCustomer = $("#btnPickCustomer");

        if (el.btnPickCustomer && el.customerPicker) {
            el.btnPickCustomer.addEventListener("click", () => {
                const isHidden = el.customerPicker.classList.contains("hidden");

                if (isHidden) {
                    el.customerPicker.classList.remove("hidden");
                    el.customerPicker.focus();
                    el.btnPickCustomer.textContent = "Cancel";
                    el.btnPickCustomer.classList.add("btn-secondary");
                } else {
                    el.customerPicker.classList.add("hidden");
                    el.btnPickCustomer.textContent = "Change";
                    el.btnPickCustomer.classList.remove("btn-secondary");
                }
            });
        }

        if (el.customerPicker) {
            el.customerPicker.addEventListener("change", async () => {
                await handleCustomerChange();

                if (el.btnPickCustomer) {
                    el.btnPickCustomer.textContent = "Change";
                    el.btnPickCustomer.classList.remove("btn-secondary");
                }
            });
        }

        // Refresh customer data button
        const btnRefreshCustomer = $("#btnRefreshCustomer");
        if (btnRefreshCustomer) {
            btnRefreshCustomer.addEventListener("click", async () => {
                // Get customer number from the field
                const customerNumber = ($("#customerNumber")?.value || "").trim();

                if (!customerNumber) {
                    alert("Please enter a Customer Number to search for.");
                    return;
                }

                const btn = btnRefreshCustomer;
                const originalHtml = btn.innerHTML;
                btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
                btn.disabled = true;

                try {
                    // Query customer by customer code/number
                    const select = "$select=" + encodeURIComponent(CUSTOMER_FIELDS.join(","));
                    const filter = "$filter=" + encodeURIComponent(`cr650_customercodes eq '${customerNumber}'`);
                    const url = `${CUSTOMER_API}?${select}&${filter}&$top=1`;

                    const response = await safeAjax({ type: "GET", url });
                    const rows = (response && (response.value || response)) || [];

                    if (rows.length === 0) {
                        alert(`Customer with number "${customerNumber}" not found.`);
                        return;
                    }

                    const refreshedCustomer = normalizeCustomer(rows[0]);
                    cache.selectedCustomer = refreshedCustomer;

                    // Update the customer in the cache list
                    const idx = cache.customers.findIndex(c => c.code === customerNumber);
                    if (idx >= 0) {
                        cache.customers[idx] = refreshedCustomer;
                    }

                    // Update UI fields
                    if (el.customerNameText) {
                        el.customerNameText.value = refreshedCustomer.code
                            ? `${refreshedCustomer.name} (${refreshedCustomer.code})`
                            : (refreshedCustomer.name || "");
                    }
                    setVal("#customerNumber", refreshedCustomer.code || "");
                    setVal("#country", refreshedCustomer.country || "");
                    setVal("#orgID", refreshedCustomer.orgId || "");
                    setVal("#COB", refreshedCustomer.cob || "");
                    setVal("#PaymentTerms", refreshedCustomer.paymentTerms || "");
                    setVal("#salesRep", refreshedCustomer.salesRep || "");
                    setVal("#destinationPort", refreshedCustomer.destinationPort || "");

                    // Update party fields based on customer data
                    const { partyType } = computePartyFieldsFromCustomer(refreshedCustomer);
                    applyShipToUIFromPartyType(partyType, refreshedCustomer);

                    // Refresh customer's available models
                    if (refreshedCustomer.id) {
                        const availableModels = await loadCustomerAvailableModels(refreshedCustomer.id);
                        cache.customerAvailableModels = availableModels;
                        renderCustomerAvailableModels(availableModels);
                    }

                    console.log("✅ Customer data refreshed successfully from customer number:", customerNumber);
                } catch (err) {
                    console.error("Failed to refresh customer data:", err);
                    alert("Failed to refresh customer data. Please try again.");
                } finally {
                    btn.innerHTML = originalHtml;
                    btn.disabled = false;
                }
            });
        }

        if ($("#addOrderBtn")) {
            $("#addOrderBtn").addEventListener("click", w.onAddOrderNumber);
        }

        const addNotifyBtn =
            d.getElementById("btnAddNotifyParty") ||
            d.getElementById("addNotifyPartyBtn");
        if (addNotifyBtn) {
            addNotifyBtn.addEventListener("click", () => {
                w.addNotifyPartyDraft();
            });
        }

        const addCustomerModelBtn = d.getElementById("btnAddCustomerModel");
        if (addCustomerModelBtn) {
            addCustomerModelBtn.addEventListener("click", () => {
                w.addCustomerModelDraft();
            });
        }

        toggleLoadingTimeInput();

        const id = getQueryParam("id");
        if (!id || !isGuid(id)) {
            alert("Invalid or missing DCL id in URL.");
            return;
        }
        cache.currentId = id;

        (function attachIdToStepLinks() {
            const guid = id;
            const nav = d.getElementById("stepIndicators");
            if (!guid || !nav) return;

            function withId(u, theId) {
                try {
                    const url = new URL(u, window.location.origin);
                    url.searchParams.set("id", theId);
                    const qs = url.searchParams.toString();
                    return url.pathname + (qs ? "?" + qs : "") + (url.hash || "");
                } catch {
                    const hasQuery = u.includes("?");
                    const hasHash = u.includes("#");
                    const [path, hash = ""] = hasHash ? u.split("#") : [u, ""];
                    const sep = hasQuery ? "&" : "?";
                    const out = `${path}${sep}id=${encodeURIComponent(theId)}`;
                    return hash ? `${out}#${hash}` : out;
                }
            }

            nav.querySelectorAll("a.step").forEach(a => {
                const href = a.getAttribute("href") || "";
                if (!href) return;
                a.setAttribute("href", withId(href, guid));
            });
        })();

        try {
            setBusy(true, "Loading DCL…");

            const rec = await loadDclById(id);
            bindDclToForm(rec);

            initCollapsibles();

            const initType = PARTY.initialType(rec);
            PARTY.applyTypeToUI(initType, rec);

            try {
                const notifyList = await loadNotifyParties(cache.currentId);
                renderNotifyPartyList(notifyList);
            } catch (errNP) {
                console.error("Failed to load notify parties", errNP);
                renderNotifyPartyList([]);
            }

            try {
                const customerModelList = await loadCustomerModels(cache.currentId);
                renderCustomerModelList(customerModelList);
            } catch (errCM) {
                console.error("Failed to load customer models", errCM);
                renderCustomerModelList([]);
            }

            // Load customer's available models if customer is selected
            if (cache.selectedCustomer && cache.selectedCustomer.id) {
                try {
                    const availableModels = await loadCustomerAvailableModels(cache.selectedCustomer.id);
                    cache.customerAvailableModels = availableModels;
                    renderCustomerAvailableModels(availableModels);
                } catch (errAM) {
                    console.error("Failed to load customer available models", errAM);
                    renderCustomerAvailableModels([]);
                }
            }

            try {
                const brandsList = await loadBrands(cache.currentId);
                renderBrandList(brandsList);
            } catch (errBrandInit) {
                console.error("Failed to load brands", errBrandInit);
                renderBrandList([]);
            }

            await ensureCurrenciesLoaded().catch(() => {
                cache.currencies = [
                    { code: "USD", name: "United States Dollar" },
                    { code: "SAR", name: "Saudi Riyal" },
                    { code: "EUR", name: "Euro" },
                    { code: "AED", name: "UAE Dirham" },
                    { code: "GBP", name: "Pound Sterling" }
                ];
            });

            populateAllCurrencySelects();

            if (rec?.cr650_currencycode) {
                setVal("#currency", rec.cr650_currencycode);
            }

            tryDefaultCurrencyByCountry();

            try {
                const shippingSelect = $("#shippingLineSelect");
                if (shippingSelect) {
                    shippingSelect.innerHTML = '<option value="">Loading shipping lines…</option>';
                    shippingSelect.disabled = true;

                    const shippingLines = await loadAllShippingLines();
                    cache.shippingLines = shippingLines;
                    populateShippingLineSelect(shippingSelect, shippingLines);

                    shippingSelect.disabled = false;

                    if (rec.cr650_shippingline) {
                        shippingSelect.value = rec.cr650_shippingline;
                    }
                }
            } catch (errShipping) {
                console.error("Failed to load shipping lines", errShipping);
                const shippingSelect = $("#shippingLineSelect");
                if (shippingSelect) {
                    shippingSelect.innerHTML = '<option value="">Failed to load shipping lines</option>';
                    shippingSelect.disabled = false;
                }
            }

            try {
                const brandSelect = $("#brandSelect");
                if (brandSelect) {
                    brandSelect.innerHTML = '<option value="">Loading brands…</option>';
                    brandSelect.disabled = true;

                    const brandsList = await loadAllBrandsList();
                    cache.brandsList = brandsList;
                    populateBrandSelect(brandSelect, brandsList);

                    brandSelect.disabled = false;
                }
            } catch (errBrands) {
                console.error("Failed to load brands list", errBrands);
                const brandSelect = $("#brandSelect");
                if (brandSelect) {
                    brandSelect.innerHTML = '<option value="">Failed to load brands</option>';
                    brandSelect.disabled = false;
                }
            }

            try {
                if (el.customerPicker) {
                    el.customerPicker.disabled = true;
                    el.customerPicker.innerHTML =
                        '<option value="">Loading customers…</option>';

                    const list = await loadAllCustomers();
                    populateCustomerSelect(list);

                    // Note: Event listener already attached earlier in the code
                }
            } catch (errCust) {
                console.error("Failed to load customers", errCust);
                if (el.customerPicker) {
                    el.customerPicker.innerHTML =
                        '<option value="">Failed to load customers</option>';
                    el.customerPicker.disabled = true;
                }
            }

            try {
                const existingDclOrders = await loadDclOrdersForCurrent();
                renderExistingDclOrders(existingDclOrders);
            } catch (errDclOrders) {
                console.error("Failed to load DCL Orders", errDclOrders);
                renderExistingDclOrders([]);
            }

            const dclCustomerNumber = rec.cr650_customernumber || "";
            if (dclCustomerNumber) {
                await reloadOrdersForCustomerNumber(dclCustomerNumber);
            } else {
                clearOrderSelectAndTable();
            }

            wireAutoSaveFields();

            // Initialize document number extraction WITH storage
            initDocumentNumberExtractionWithStorage();

            if (cache.currentDcl?.cr650_currencycode) {
                const cur = cache.currentDcl.cr650_currencycode;
                const curSel = document.getElementById("currency");

                if (curSel) {
                    curSel.value = cur;
                    curSel.dispatchEvent(new Event("change", { bubbles: true }));
                    console.log("✅ Final forced currency restore:", cur);
                }
            }

            lockFormIfSubmitted();

        } catch (err) {
            console.error("❌ Failed to load DCL:", err);
            console.error("Error details:", {
                message: err.message,
                stack: err.stack,
                dclId: cache.currentId
            });

            let errorMsg = "Unable to load DCL record.\n\n";

            if (err.message) {
                errorMsg += "Error: " + err.message + "\n\n";
            }

            errorMsg += "Please check:\n";
            errorMsg += "• Do you have permission to view this DCL?\n";
            errorMsg += "• Is the DCL ID correct?\n";
            errorMsg += "• Check browser console for details (F12)";

            alert(errorMsg);
        } finally {
            setBusy(false);
        }
    });

})(window, document);