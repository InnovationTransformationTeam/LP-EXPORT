(function (w, d) {
    "use strict";

    /************************************************************
     * 0. GLOBAL STATE + CONSTANTS / ENDPOINTS
     ************************************************************/
    w.currentDclId = w.currentDclId || null;

    const DCL_API = "/_api/cr650_dcl_masters";
    const CUSTOMER_API = "/_api/cr650_updated_dcl_customers";
    const OUTSTANDING_API = "/_api/cr650_dcl_outstanding_reports";
    const NOTIFY_API = "/_api/cr650_dcl_notify_parties";
    const TERMS_API = "/_api/cr650_dcl_terms_conditionses";
    const SHIPPING_LINES_API = "/_api/cr650_dcl_shippinglines";
    const BRANDS_API = "/_api/cr650_dcl_brands";
    const BRANDS_LIST_API = "/_api/cr650_dcl_brands_lists";
    const DCL_ORDERS_API = "/_api/cr650_dcl_orders";
    const CUSTOMER_MODEL_API = "/_api/cr650_dcl_customer_models";

    const CUSTOMER_MODEL_FIELDS = [
        "cr650_dcl_customer_modelid",
        "cr650_model_name",
        "cr650_info",
        "_cr650_dcl_id_value",
        "_cr650_cutomer_id_value",
        "cr650_is_related_to_a_customer",
        "createdon"
    ];

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
    "cr650_notifyparty2",
    "cr650_currency"
    ];




    const ORDER_FIELDS = [
        "cr650_customer_number",
        "cr650_order_no",
        "cr650_source_order_number",
        "cr650_cust_po_number",
        "createdon"
    ];

    const CURRENCY_SOURCES = [
        "https://openexchangerates.org/api/currencies.json",
        "https://api.frankfurter.app/currencies"
    ];
    const CURRENCY_CACHE_KEY = "dcl_currencies_v2";
    const CURRENCY_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h

    /************************************************************
     * 1. RUNTIME CACHE
     ************************************************************/
    const cache = {
        customers: [],
        selectedCustomer: null,
        currencies: [],
        shippingLines: [],
        customerModels: [],
        pendingDCLModels: [], // Models to be created when DCL is saved
        brandsList: []
    };

    const $ = (sel, root = d) => root.querySelector(sel);
    const $$ = (sel, root = d) => Array.from(root.querySelectorAll(sel));

    const el = {};


    /************************************************************
 * COUNTRY CODE MAPPING (ISO 3166-1 Alpha-3)
 ************************************************************/
    const COUNTRY_CODES = {
        "AFGHANISTAN": "AFG",
        "ALBANIA": "ALB",
        "ALGERIA": "DZA",
        "AMERICAN SAMOA": "ASM",
        "ANDORRA": "AND",
        "ANGOLA": "AGO",
        "ANTIGUA AND BARBUDA": "ATG",
        "ARGENTINA": "ARG",
        "ARMENIA": "ARM",
        "AUSTRALIA": "AUS",
        "AUSTRIA": "AUT",
        "AZERBAIJAN": "AZE",
        "BAHAMAS": "BHS",
        "BAHRAIN": "BHR",
        "BANGLADESH": "BGD",
        "BARBADOS": "BRB",
        "BELARUS": "BLR",
        "BELGIUM": "BEL",
        "BELIZE": "BLZ",
        "BENIN": "BEN",
        "BHUTAN": "BTN",
        "BOLIVIA": "BOL",
        "BOSNIA AND HERZEGOVINA": "BIH",
        "BOTSWANA": "BWA",
        "BRAZIL": "BRA",
        "BRUNEI": "BRN",
        "BULGARIA": "BGR",
        "BURKINA FASO": "BFA",
        "BURUNDI": "BDI",
        "CAMBODIA": "KHM",
        "CAMEROON": "CMR",
        "CANADA": "CAN",
        "CAPE VERDE": "CPV",
        "CENTRAL AFRICAN REPUBLIC": "CAF",
        "CHAD": "TCD",
        "CHILE": "CHL",
        "CHINA": "CHN",
        "COLOMBIA": "COL",
        "COMOROS": "COM",
        "CONGO": "COG",
        "COSTA RICA": "CRI",
        "CROATIA": "HRV",
        "CUBA": "CUB",
        "CYPRUS": "CYP",
        "CZECH REPUBLIC": "CZE",
        "DENMARK": "DNK",
        "DJIBOUTI": "DJI",
        "DOMINICA": "DMA",
        "DOMINICAN REPUBLIC": "DOM",
        "ECUADOR": "ECU",
        "EGYPT": "EGY",
        "EL SALVADOR": "SLV",
        "EQUATORIAL GUINEA": "GNQ",
        "ERITREA": "ERI",
        "ESTONIA": "EST",
        "ETHIOPIA": "ETH",
        "FIJI": "FJI",
        "FINLAND": "FIN",
        "FRANCE": "FRA",
        "GABON": "GAB",
        "GAMBIA": "GMB",
        "GEORGIA": "GEO",
        "GERMANY": "DEU",
        "GHANA": "GHA",
        "GREECE": "GRC",
        "GRENADA": "GRD",
        "GUATEMALA": "GTM",
        "GUINEA": "GIN",
        "GUINEA-BISSAU": "GNB",
        "GUYANA": "GUY",
        "HAITI": "HTI",
        "HONDURAS": "HND",
        "HUNGARY": "HUN",
        "ICELAND": "ISL",
        "INDIA": "IND",
        "INDONESIA": "IDN",
        "IRAN": "IRN",
        "IRAQ": "IRQ",
        "IRELAND": "IRL",
        "ISRAEL": "ISR",
        "ITALY": "ITA",
        "JAMAICA": "JAM",
        "JAPAN": "JPN",
        "JORDAN": "JOR",
        "KAZAKHSTAN": "KAZ",
        "KENYA": "KEN",
        "KIRIBATI": "KIR",
        "KOREA, NORTH": "PRK",
        "KOREA, SOUTH": "KOR",
        "KUWAIT": "KWT",
        "KYRGYZSTAN": "KGZ",
        "LAOS": "LAO",
        "LATVIA": "LVA",
        "LEBANON": "LBN",
        "LESOTHO": "LSO",
        "LIBERIA": "LBR",
        "LIBYA": "LBY",
        "LIECHTENSTEIN": "LIE",
        "LITHUANIA": "LTU",
        "LUXEMBOURG": "LUX",
        "MADAGASCAR": "MDG",
        "MALAWI": "MWI",
        "MALAYSIA": "MYS",
        "MALDIVES": "MDV",
        "MALI": "MLI",
        "MALTA": "MLT",
        "MARSHALL ISLANDS": "MHL",
        "MAURITANIA": "MRT",
        "MAURITIUS": "MUS",
        "MEXICO": "MEX",
        "MICRONESIA": "FSM",
        "MOLDOVA": "MDA",
        "MONACO": "MCO",
        "MONGOLIA": "MNG",
        "MONTENEGRO": "MNE",
        "MOROCCO": "MAR",
        "MOZAMBIQUE": "MOZ",
        "MYANMAR": "MMR",
        "NAMIBIA": "NAM",
        "NAURU": "NRU",
        "NEPAL": "NPL",
        "NETHERLANDS": "NLD",
        "NEW ZEALAND": "NZL",
        "NICARAGUA": "NIC",
        "NIGER": "NER",
        "NIGERIA": "NGA",
        "NORTH MACEDONIA": "MKD",
        "NORWAY": "NOR",
        "OMAN": "OMN",
        "PAKISTAN": "PAK",
        "PALAU": "PLW",
        "PANAMA": "PAN",
        "PAPUA NEW GUINEA": "PNG",
        "PARAGUAY": "PRY",
        "PERU": "PER",
        "PHILIPPINES": "PHL",
        "POLAND": "POL",
        "PORTUGAL": "PRT",
        "QATAR": "QAT",
        "ROMANIA": "ROU",
        "RUSSIA": "RUS",
        "RWANDA": "RWA",
        "SAINT KITTS AND NEVIS": "KNA",
        "SAINT LUCIA": "LCA",
        "SAINT VINCENT AND THE GRENADINES": "VCT",
        "SAMOA": "WSM",
        "SAN MARINO": "SMR",
        "SAO TOME AND PRINCIPE": "STP",
        "SAUDI ARABIA": "SAU",
        "SENEGAL": "SEN",
        "SERBIA": "SRB",
        "SEYCHELLES": "SYC",
        "SIERRA LEONE": "SLE",
        "SINGAPORE": "SGP",
        "SLOVAKIA": "SVK",
        "SLOVENIA": "SVN",
        "SOLOMON ISLANDS": "SLB",
        "SOMALIA": "SOM",
        "SOUTH AFRICA": "ZAF",
        "SOUTH SUDAN": "SSD",
        "SPAIN": "ESP",
        "SRI LANKA": "LKA",
        "SUDAN": "SDN",
        "SURINAME": "SUR",
        "SWEDEN": "SWE",
        "SWITZERLAND": "CHE",
        "SYRIA": "SYR",
        "TAIWAN": "TWN",
        "TAJIKISTAN": "TJK",
        "TANZANIA": "TZA",
        "THAILAND": "THA",
        "TIMOR-LESTE": "TLS",
        "TOGO": "TGO",
        "TONGA": "TON",
        "TRINIDAD AND TOBAGO": "TTO",
        "TUNISIA": "TUN",
        "TURKEY": "TUR",
        "TURKMENISTAN": "TKM",
        "TUVALU": "TUV",
        "UGANDA": "UGA",
        "UKRAINE": "UKR",
        "UNITED ARAB EMIRATES": "ARE",
        "UAE": "ARE",
        "UNITED KINGDOM": "GBR",
        "UK": "GBR",
        "UNITED STATES": "USA",
        "USA": "USA",
        "URUGUAY": "URY",
        "UZBEKISTAN": "UZB",
        "VANUATU": "VUT",
        "VATICAN CITY": "VAT",
        "VENEZUELA": "VEN",
        "VIETNAM": "VNM",
        "YEMEN": "YEM",
        "ZAMBIA": "ZMB",
        "ZIMBABWE": "ZWE"
    };

    /**
     * Get country code from country name
     */
    function getCountryCode(countryName) {
        if (!countryName) return "UNK"; // Unknown

        const normalized = countryName.trim().toUpperCase();

        // Direct match
        if (COUNTRY_CODES[normalized]) {
            return COUNTRY_CODES[normalized];
        }

        // Partial match (for cases like "United Arab Emirates, Dubai")
        for (const [key, code] of Object.entries(COUNTRY_CODES)) {
            if (normalized.includes(key) || key.includes(normalized)) {
                return code;
            }
        }

        return "UNK"; // Unknown country
    }

    /************************************************************
     * 2. UTILITIES
     ************************************************************/
    function escapeHtml(str) {
        return String(str ?? "").replace(/[&<>"']/g, m => ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#39;"
        }[m]));
    }

    function asISO(dateVal) {
        if (!dateVal) return null;
        const dt = new Date(dateVal + "T00:00:00");
        if (isNaN(dt)) return null;
        return dt.toISOString();
    }

    function setVal(sel, v) {
        const node = $(sel);
        if (!node) return;
        const tag = node.tagName.toLowerCase();
        if (tag === "input" || tag === "select" || tag === "textarea") {
            node.value = v ?? "";
        } else {
            node.textContent = v ?? "";
        }
    }

    function getSelectedCustomerNameOrEmpty() {
        const picker = $("#customerPicker");
        if (!picker) return "";
        const idx = picker.options[picker.selectedIndex]?.dataset.idx;
        if (idx == null) return "";
        const cust = cache.customers[idx];
        if (!cust) return "";
        return cust.name || "";
    }


    /************************************************************
 * COMPANY TYPE COLOR HANDLER
 ************************************************************/
    /************************************************************
 * 15.5. COMPANY TYPE COLOR HANDLER
 ************************************************************/
    function updateCompanyTypeColor() {
        const select = $("#companyType");
        const icon = $("#companyTypeIcon"); // Optional icon element

        if (!select) return;

        // Remove all color classes
        select.classList.remove("petrolube-selected", "technolube-selected");

        // Hide icon if it exists
        if (icon) {
            icon.style.display = "none";
        }

        const value = select.value;

        if (value === "1") {
            // Petrolube - Green
            select.classList.add("petrolube-selected");
            if (icon) {
                icon.textContent = "✓";
                icon.style.color = "#16a34a";
                icon.style.display = "block";
            }
        } else if (value === "0") {
            // Technolube - Red
            select.classList.add("technolube-selected");
            if (icon) {
                icon.textContent = "✓";
                icon.style.color = "#dc2626";
                icon.style.display = "block";
            }
        }
    }

    w.updateCompanyTypeColor = updateCompanyTypeColor; // Expose globally if needed

    /************************************************************
     * 3. TOP LOADER + safeAjax
     ************************************************************/
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
            forceClear() { refs = 0; hide(); }
        };
    })();
    w.TopLoader = TopLoader;

    function safeAjax(options) {
        return new Promise((resolve, reject) => {
            options = options || {};
            options.type = options.type || "GET";
            options.headers = options.headers || {};

            // we will keep this for clarity even though we don't always apply it directly:
            options.headers.Accept = options.headers.Accept
                || "application/json;odata.metadata=minimal";
            options.contentType = options.contentType
                || "application/json; charset=utf-8";

            async function go(token) {
                if (token) {
                    options.headers["__RequestVerificationToken"] = token;
                }
                // ensure content type header if caller didn't set one and we're sending data
                if (options.data && !options.headers["Content-Type"]) {
                    options.headers["Content-Type"] = options.contentType;
                }

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
                        try {
                            reject(JSON.parse(text));
                        } catch {
                            reject(new Error(text || ("HTTP " + res.status)));
                        }
                        return;
                    }

                    if (!text) {
                        resolve({});
                        return;
                    }

                    try {
                        resolve(JSON.parse(text));
                    } catch {
                        reject(new Error("Response was not valid JSON."));
                    }
                } catch (e) {
                    reject(e);
                } finally {
                    TopLoader.end();
                }
            }

            if (w.shell && typeof shell.getTokenDeferred === "function") {
                shell.getTokenDeferred()
                    .done(go)
                    .fail(() => reject({
                        message: "Token API unavailable (user not authenticated?)"
                    }));
            } else {
                go(null);
            }
        });
    }
    w.safeAjax = safeAjax;

    w.addEventListener("beforeunload", () => {
        const bar = d.getElementById("top-loader");
        if (bar) bar.classList.remove("hidden");
        const wiz = d.getElementById("dclWizard");
        if (wiz) wiz.classList.add("blur-while-loading");
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

    /************************************************************
     * 4. PATCH (LIVE SAVE)
     ************************************************************/
    async function getAntiForgeryTokenMaybe() {
        if (w.shell && typeof shell.getTokenDeferred === "function") {
            return new Promise((resolve, reject) => {
                shell.getTokenDeferred()
                    .done(resolve)
                    .fail(() => reject("Token API unavailable"));
            }).catch(() => null);
        }
        return null;
    }

    async function patchDclField(dclId, partialBodyObj) {
        if (!dclId) {
            console.warn("No DCL ID yet; skipping PATCH.", partialBodyObj);
            return;
        }

        const url = `${DCL_API}(${dclId})`;

        const headers = {
            "Accept": "application/json;odata.metadata=minimal",
            "Content-Type": "application/json; charset=utf-8",
            "If-Match": "*"
        };

        const token = await getAntiForgeryTokenMaybe();
        if (token) {
            headers["__RequestVerificationToken"] = token;
        }

        TopLoader.begin();
        try {
            const res = await fetch(url, {
                method: "PATCH",
                headers,
                body: JSON.stringify(partialBodyObj)
            });

            if (!res.ok) {
                const txt = await res.text();
                console.error("PATCH failed", res.status, txt);
                alert("Failed to save changes to server.");
            }
        } catch (err) {
            console.error("PATCH error", err);
            alert("Network error while saving.");
        } finally {
            TopLoader.end();
        }
    }

    /************************************************************
     * 5. CURRENCIES
     ************************************************************/
    function fetchWithTimeout(url, ms = 8000) {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), ms);

        return fetch(url, { signal: ctrl.signal })
            .then(res => {
                return res.text().then(txt => {
                    let data = null;
                    try { data = txt ? JSON.parse(txt) : null; } catch { }
                    return { ok: res.ok, data };
                });
            })
            .catch(() => ({ ok: false, data: null }))
            .finally(() => clearTimeout(t));
    }

    const CURRENCY_CACHE_TTL = CURRENCY_CACHE_TTL_MS;
    function currenciesFromCache() {
        try {
            const cached = JSON.parse(
                w.localStorage.getItem(CURRENCY_CACHE_KEY) || "null"
            );
            if (
                cached &&
                Array.isArray(cached.items) &&
                (Date.now() - cached.savedAt) < CURRENCY_CACHE_TTL
            ) {
                return cached.items;
            }
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
        if (cached) {
            cache.currencies = cached;
            return;
        }

        const results = await Promise.allSettled(
            CURRENCY_SOURCES.map(u => fetchWithTimeout(u))
        );

        const dicts = results
            .filter(r => r.status === "fulfilled" && r.value.ok)
            .map(r => r.value.data)
            .filter(Boolean);

        let union = {};
        dicts.forEach(dct => {
            const norm = normalizeCurrencyMap(dct);
            Object.keys(norm).forEach(c => {
                if (!union[c]) union[c] = norm[c];
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
        const previous = selectEl.value;

        const opts = ['<option value="">Select Currency</option>']
            .concat(
                cache.currencies.map(
                    c => `<option value="${c.code}">${escapeHtml(`${c.code} - ${c.name}`)}</option>`
                )
            );

        selectEl.innerHTML = opts.join("");

        if (previous) {
            selectEl.value = previous;
        }
    }

    function populateAllCurrencySelects() {
        $$(".currency-select").forEach(populateCurrencySelect);
    }

    function tryDefaultCurrencyByCountry() {
        const curSel = $("#currency");
        if (!curSel || curSel.value) return;

        const countryVal = ($("#country")?.value || "").toUpperCase();

        function prefer(code) {
            if (curSel.querySelector(`option[value='${code}']`)) {
                curSel.value = code;
            }
        }

        if (countryVal.includes("SAUDI") || countryVal === "SA") {
            prefer("SAR");
        } else if (
            countryVal.includes("UAE") ||
            countryVal.includes("EMIRATES") ||
            countryVal.includes("UNITED ARAB") ||
            countryVal === "AE"
        ) {
            prefer("AED");
        } else if (
            countryVal.includes("UNITED STATES") ||
            countryVal.includes("USA") ||
            countryVal === "US"
        ) {
            prefer("USD");
        } else if (
            countryVal.includes("UNITED KINGDOM") ||
            countryVal.includes("UK") ||
            countryVal.includes("BRITAIN")
        ) {
            prefer("GBP");
        } else if (
            countryVal.includes("EUROPE") ||
            countryVal.includes("GERMANY") ||
            countryVal === "DE" ||
            countryVal === "EU"
        ) {
            prefer("EUR");
        }
    }

    /************************************************************
 * 5.5. SHIPPING LINES
 ************************************************************/
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

        // Add the "Add new" option at the end
        const addOpt = document.createElement("option");
        addOpt.value = "__add_new__";
        addOpt.textContent = "+ Add new shipping line";
        selectEl.appendChild(addOpt);

        // Restore previous selection if it still exists
        if (currentValue && selectEl.querySelector(`option[value="${currentValue}"]`)) {
            selectEl.value = currentValue;
        }
    }

    /************************************************************
     * 5b. BRANDS LIST (Lookup Table)
     ************************************************************/
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

        // Add the "Add new" option at the end
        const addOpt = document.createElement("option");
        addOpt.value = "__add_new__";
        addOpt.textContent = "+ Add new brand";
        selectEl.appendChild(addOpt);

        // Restore previous selection if it still exists
        if (currentValue && selectEl.querySelector(`option[value="${currentValue}"]`)) {
            selectEl.value = currentValue;
        }
    }

    /************************************************************
     * 6. CUSTOMERS + ORDERS
     ************************************************************/
    function normalizeCustomer(r) {
        const g = (x) => (x == null ? "" : x);

        return {
            id: g(r.cr650_updated_dcl_customerid),   // PK
            name: g(r.cr650_customername),
            code: g(r.cr650_customercodes),

            country: g(r.cr650_country),
            countryCode: g(r.cr650_country_1),

            orgId: g(r.cr650_organizationid),        // ✅ Technolube (UAE)
            cob: g(r.cr650_cob),
            paymentTerms: g(r.cr650_paymentterms),
            salesRep: g(r.cr650_salesrepresentativename),
            destinationPort: g(r.cr650_destinationport),
            currency: g(r.cr650_currency),

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

                // ✅ Correct sort
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
            const label = c.code
                ? `${c.name} (${c.code})`
                : (c.name || "(Unnamed)");
            const opt = d.createElement("option");
            opt.value = c.id || "";
            opt.textContent = label;
            opt.dataset.idx = String(i);
            sel.appendChild(opt);
        });

        sel.disabled = false;
    }

    function trySelectCustomerOnUI(cust) {
        // (Optional field, may not exist in simplified HTML)
        const cnText = $("#customerNameText");
        if (cnText) {
            cnText.value = cust.code
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

        // Auto-fill currency from customer
        if (cust.currency) {
            setVal("#currency", cust.currency);
        }
    }

    // =========================================================================
    // CUSTOMER MODELS FUNCTIONS
    // =========================================================================

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

    function buildCustomerModelQueryUrl(customerId) {
        const select = "$select=" + encodeURIComponent(CUSTOMER_MODEL_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_cutomer_id_value eq ${customerId} and cr650_is_related_to_a_customer eq true`);
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${CUSTOMER_MODEL_API}?${select}&${filter}&${orderby}`;
    }

    async function loadCustomerModelsForCustomer(customerId) {
        if (!customerId) return [];
        const records = [];

        async function loop(nextUrl) {
            const resp = await fetch(nextUrl, {
                headers: {
                    'Accept': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });
            if (!resp.ok) return;
            const data = await resp.json();
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildCustomerModelQueryUrl(customerId));
        return records.map(normalizeCustomerModel);
    }

    async function createCustomerModelForDCL(dclId, modelName, info) {
        if (!dclId) throw new Error("No DCL ID provided");

        const bodyObj = {
            cr650_model_name: modelName,
            cr650_info: info || "",
            cr650_is_related_to_a_customer: false,
            "cr650_dcl_id@odata.bind": "/cr650_dcl_masters(" + dclId + ")"
        };

        await safeAjax({
            type: "POST",
            url: CUSTOMER_MODEL_API,
            contentType: 'application/json',
            data: JSON.stringify(bodyObj)
        });

        return { ok: true };
    }

    async function deleteCustomerModelFromDCL(modelId) {
        const url = `${CUSTOMER_MODEL_API}(${encodeURIComponent(modelId)})`;
        await safeAjax({
            type: "DELETE",
            url,
            headers: { "If-Match": "*" }
        });
        return { ok: true };
    }

    function escapeHtmlForModel(str) {
        if (!str) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function renderCustomerModelsSection(models) {
        const container = $("#customerModelsContainer");
        if (!container) return;

        if (!models || models.length === 0) {
            container.innerHTML = `
                <div style="text-align: center; padding: 20px; background: #f9f9f9; border: 1px dashed #ccc; border-radius: 8px;">
                    <i class="fas fa-inbox" style="font-size: 2rem; color: #999; margin-bottom: 8px;"></i>
                    <p style="color: #666; margin: 0;">No customer models available for this customer.</p>
                    <p style="color: #999; font-size: 0.85rem; margin-top: 4px;">Select a customer to see their models, or add a new model below.</p>
                </div>
            `;
            return;
        }

        container.innerHTML = models.map((model, index) => `
            <div class="customer-model-card" data-model-id="${model.id}" data-index="${index}" style="
                background: #fff;
                border: 1px solid #e0e0e0;
                border-left: 4px solid var(--primary-green, #28a745);
                border-radius: 8px;
                padding: 12px 16px;
                margin-bottom: 10px;
            ">
                <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                    <div style="flex: 1;">
                        <strong style="font-size: 1rem; color: #333;">${escapeHtmlForModel(model.modelName)}</strong>
                        <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #666; white-space: pre-wrap;">${model.info ? escapeHtmlForModel(model.info) : '<em style="color: #999;">No additional info</em>'}</p>
                    </div>
                    <div style="display: flex; gap: 6px; margin-left: 12px;">
                        <button type="button" class="btn-use-model" data-model-name="${escapeHtmlForModel(model.modelName)}" data-model-info="${escapeHtmlForModel(model.info)}" style="
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
    }

    async function handleAddModelToDCL() {
        const nameInput = $("#newModelName");
        const infoInput = $("#newModelInfo");

        const modelName = (nameInput?.value || "").trim();
        const modelInfo = (infoInput?.value || "").trim();

        if (!modelName) {
            alert("Please enter a model name.");
            if (nameInput) nameInput.focus();
            return;
        }

        // If DCL doesn't exist yet, store model locally
        if (!w.currentDclId) {
            cache.pendingDCLModels.push({
                id: 'pending-' + Date.now(),
                modelName: modelName,
                info: modelInfo,
                isPending: true
            });

            // Clear form
            if (nameInput) nameInput.value = "";
            if (infoInput) infoInput.value = "";

            // Render pending models
            renderDCLModelsSection(cache.pendingDCLModels);
            return;
        }

        // DCL exists - create model directly
        try {
            await createCustomerModelForDCL(w.currentDclId, modelName, modelInfo);

            // Clear form
            if (nameInput) nameInput.value = "";
            if (infoInput) infoInput.value = "";

            // Refresh DCL models list
            await refreshDCLModels();
        } catch (err) {
            console.error("Failed to add model:", err);
            alert("Failed to add model. Please try again.");
        }
    }

    async function refreshDCLModels() {
        if (!w.currentDclId) return;

        try {
            const dclModels = await loadDCLModels(w.currentDclId);
            renderDCLModelsSection(dclModels);
        } catch (err) {
            console.error("Failed to refresh DCL models:", err);
        }
    }

    function buildDCLModelQueryUrl(dclId) {
        const select = "$select=" + encodeURIComponent(CUSTOMER_MODEL_FIELDS.join(","));
        const filter = "$filter=" + encodeURIComponent(`_cr650_dcl_id_value eq ${dclId}`);
        const orderby = "$orderby=" + encodeURIComponent("createdon asc");
        return `${CUSTOMER_MODEL_API}?${select}&${filter}&${orderby}`;
    }

    async function loadDCLModels(dclId) {
        if (!dclId) return [];
        const records = [];

        async function loop(nextUrl) {
            const resp = await fetch(nextUrl, {
                headers: {
                    'Accept': 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0'
                }
            });
            if (!resp.ok) return;
            const data = await resp.json();
            const rows = (data && (data.value || data)) || [];
            rows.forEach(r => records.push(r));
            if (data && data["@odata.nextLink"]) {
                return loop(data["@odata.nextLink"]);
            }
        }

        await loop(buildDCLModelQueryUrl(dclId));
        return records.map(normalizeCustomerModel);
    }

    function renderDCLModelsSection(models) {
        const container = $("#dclModelsContainer");
        if (!container) return;

        if (!models || models.length === 0) {
            container.innerHTML = `
                <div style="text-align: center; padding: 16px; background: #fff8e1; border: 1px dashed #ffcc80; border-radius: 8px;">
                    <p style="color: #f57c00; margin: 0; font-size: 0.9rem;"><i class="fas fa-info-circle" style="margin-right: 6px;"></i>No models added yet. Use models from above or add a new one.</p>
                </div>
            `;
            return;
        }

        container.innerHTML = models.map(model => {
            const isPending = model.isPending || (model.id && model.id.toString().startsWith('pending-'));
            const pendingBadge = isPending
                ? '<span style="background: #fff3cd; color: #856404; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; margin-left: 8px;">Pending</span>'
                : '';

            return `
            <div class="dcl-model-card" data-model-id="${model.id}" style="
                background: #fff;
                border: 1px solid ${isPending ? '#ffe082' : '#c8e6c9'};
                border-left: 4px solid ${isPending ? '#ffc107' : '#4caf50'};
                border-radius: 8px;
                padding: 12px 16px;
            ">
                <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                    <div style="flex: 1;">
                        <strong style="font-size: 1rem; color: #333;">${escapeHtmlForModel(model.modelName)}${pendingBadge}</strong>
                        <p style="margin: 4px 0 0 0; font-size: 0.875rem; color: #666; white-space: pre-wrap;">${model.info ? escapeHtmlForModel(model.info) : '<em style="color: #999;">No additional info</em>'}</p>
                    </div>
                    <button type="button" class="btn-delete-dcl-model" data-model-id="${model.id}" style="
                        padding: 4px 10px;
                        background: #ffebee;
                        color: #c62828;
                        border: 1px solid #ef9a9a;
                        border-radius: 4px;
                        font-size: 0.75rem;
                        cursor: pointer;
                        margin-left: 12px;
                    "><i class="fas fa-trash"></i> Remove</button>
                </div>
            </div>
        `}).join('');
    }

    // =========================================================================
    // END CUSTOMER MODELS FUNCTIONS
    // =========================================================================

    function normalizeOrderRecord(r) {
        const s = (v) => (v == null ? "" : String(v).trim());

        return {
            orderValue: s(r.cr650_order_no),
            label: s(r.cr650_order_no),
            custPoNumber: s(r.cr650_cust_po_number)
        };
    }


    function buildOrderQueryUrl(customerNumber) {
        const select = "$select=" + encodeURIComponent([
            "cr650_customer_number",
            "cr650_order_no",
            "cr650_cust_po_number",
            "createdon"
        ].join(","));

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
            const key = (o.orderValue || "").trim();
            if (!key) continue;

            if (!seen.has(key)) {
                seen.add(key);
                out.push(o);   // ✅ ONE ROW PER ORDER HEADER
            }
        }

        return out;
    }

    function setOrderSelectLoading(isLoading, msg = "Loading orders…") {
        const sel = $("#orderNumberSelect");
        if (!sel) return;
        if (isLoading) {
            sel.innerHTML = `<option value="">${msg}</option>`;
            sel.disabled = true;
        } else {
            sel.disabled = false;
        }
    }

    function populateOrderSelect(orders) {
        const sel = $("#orderNumberSelect");
        if (!sel) return;

        const deduped = dedupeOrders(orders);

        // ✅ UI FEEDBACK WHEN NO ORDERS
        if (!deduped.length) {
            sel.innerHTML = '<option value="">No outstanding orders for this customer</option>';
            sel.disabled = true;
            return;
        }

        sel.disabled = false;
        sel.innerHTML = '<option value="">Select Order Number</option>';

        deduped.forEach(o => {
            const opt = document.createElement("option");
            opt.value = o.orderValue;
            opt.textContent = o.label;
            opt.dataset.custPoNumber = o.custPoNumber || "";
            sel.appendChild(opt);
        });
    }


    function clearOrderSelectAndTable() {
        const sel = $("#orderNumberSelect");
        const tbody = $("#orderNumbersTable tbody");
        const tableSection = $("#orderNumbersTableSection");

        if (sel) {
            sel.innerHTML = '<option value="">Select Order Number</option>';
            sel.disabled = false;
        }
        if (tbody) tbody.innerHTML = "";
        if (tableSection) tableSection.style.display = "none";
    }

    async function reloadOrdersForNewCustomerNumber(newCustNum) {
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
            const orderSel = $("#orderNumberSelect");
            if (orderSel) {
                orderSel.innerHTML = '<option value="">Failed to load orders</option>';
                orderSel.disabled = true;
            }
        } finally {
            setOrderSelectLoading(false);
        }
    }

    /************************************************************
     * 7. PARTY PREFILL
     ************************************************************/
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
            patch.partyLabel = "Ship To";
            patch.shiptolocation = cust.shipTo || "";
            patch.billto = cust.billTo || "";
            patch.consignee = "";
            patch.applicant = "";
        } else if (partyType === "consignee") {
            patch.partyLabel = "Consignee";
            patch.consignee = cust.consignee || "";
            patch.shiptolocation = "";
            patch.billto = "";
            patch.applicant = "";
        } else {
            patch.partyLabel = "Applicant";
            patch.applicant = cust.primaryAddress || "";
            patch.shiptolocation = "";
            patch.billto = "";
            patch.consignee = "";
        }

        return { patch, partyType };
    }

    function fillPartyUIFromRecord(partyType, record) {
        const billBox = $("#billToContainer");
        const billTA = $("#billTo");
        const mainTA = $("#shipTo");
        const partySel = $("#shipToLabel");

        if (partySel) {
            partySel.value = partyType;
        }

        const shipToVal = record?.shipTo || "";
        const billToVal = record?.billTo || "";
        const consigneeVal = record?.consignee || "";
        const applicantVal = record?.primaryAddress || "";

        if (partyType === "shipto") {
            if (mainTA) mainTA.value = shipToVal;
            if (billTA) billTA.value = billToVal;
            if (billBox) billBox.style.display = "";
            if (billTA) billTA.removeAttribute("disabled");
        } else if (partyType === "consignee") {
            if (mainTA) mainTA.value = consigneeVal;
            if (billBox) billBox.style.display = "none";
            if (billTA) billTA.setAttribute("disabled", "disabled");
        } else { // applicant
            if (mainTA) mainTA.value = applicantVal;
            if (billBox) billBox.style.display = "none";
            if (billTA) billTA.setAttribute("disabled", "disabled");
        }
    }

    function applyShipToUIFromPartyType(partyType, custRecord) {
        fillPartyUIFromRecord(partyType, custRecord);
    }

    /************************************************************
     * 8. INLINE UI BUILDERS
     ************************************************************/
    function buildNotifyPartyRow(textVal = "") {
        const row = d.createElement("div");
        row.className = "notify-card";

        row.innerHTML = `
            <textarea
                class="notify-text"
                rows="4"
                placeholder="Enter Notify Party details (name, address, contact)">${escapeHtml(textVal)}</textarea>

            <div class="notify-actions">
                <div class="notify-buttons">
                    <button
                        type="button"
                        class="notify-delete-btn"
                    >Remove</button>
                </div>
            </div>
        `;
        return row;
    }

    function appendNotifyPartyRow() {
        const container = $("#notifyPartyContainer");
        if (!container) return;
        const rowEl = buildNotifyPartyRow("");
        container.appendChild(rowEl);
    }

    function buildTermsConditionItem(textVal = "", printCI = false, printPL = false, index = 1) {
        const wrap = d.createElement("div");
        wrap.className = "terms-condition-item";
        wrap.style.cssText =
            "position:relative; margin-bottom:1rem; border:1px solid #e5e7eb; padding:1rem; border-radius:6px; background:#fff;";

        wrap.innerHTML = `
            <div class="form-group">
                <label for="termsComments${index}">Terms &amp; Conditions</label>
                <textarea
                    id="termsComments${index}"
                    class="terms-textarea"
                    rows="4"
                    style="
                        width:100%;
                        padding:8px;
                        font-size:14px;
                        border:1px solid #ccc;
                        border-radius:4px;
                        min-height:120px;
                        resize:vertical;
                    "
                >${escapeHtml(textVal)}</textarea>
            </div>

            <div class="form-group">
                <span>Print on:</span>
                <div style="display:flex; gap:1rem; margin-top:0.5rem;">
                    <label style="display:flex; align-items:center; gap:0.5rem;">
                        <input
                            type="checkbox"
                            class="tc-print-ci"
                            ${printCI ? "checked" : ""}>
                        <span>Commercial Invoice (CI)</span>
                    </label>

                    <label style="display:flex; align-items:center; gap:0.5rem;">
                        <input
                            type="checkbox"
                            class="tc-print-pl"
                            ${printPL ? "checked" : ""}>
                        <span>Packing List (PL)</span>
                    </label>
                </div>
            </div>

            <div class="notify-actions" style="margin-top:.75rem;">
                <div class="notify-buttons">
                    <button
                        type="button"
                        class="tc-delete-btn"
                        style="
                            background-color:#dc2626;
                            border-color:#dc2626;
                            color:#fff;
                            font-size:.8rem;
                            line-height:1.2;
                            font-weight:600;
                            padding:.5rem .75rem;
                            border-radius:.5rem;
                            border:1px solid transparent;
                            cursor:pointer;
                            min-width:3.5rem;
                            text-align:center;
                            box-shadow:0 1px 2px rgba(0,0,0,.05);
                        "
                    >Remove</button>
                </div>
            </div>
        `;
        return wrap;
    }

    function appendTermsConditionRow() {
        const container = $("#termsConditionsContainer");
        if (!container) return;
        const idx = container.children.length + 1;
        const rowEl = buildTermsConditionItem("", false, false, idx);
        container.appendChild(rowEl);
    }

    function buildBrandChip(brandName) {
        const chip = d.createElement("span");
        chip.className = "brand-chip";
        chip.dataset.value = brandName;
        chip.innerHTML = `
            <span>${escapeHtml(brandName)}</span>
            <button
                type="button"
                class="brand-remove-btn"
                aria-label="Remove brand"
            >×</button>
        `;
        return chip;
    }

    async function addBrandChipFromSelect() {
        const select = $("#brandSelect");
        const wrap = $("#brandList");
        if (!select || !wrap) return;

        let value = select.value;
        let text = select.options[select.selectedIndex]?.textContent?.trim() || "";

        // Handle "Add new brand" option
        if (value === "__add_new__") {
            const newBrandName = prompt("Enter new brand name:");

            if (newBrandName && newBrandName.trim()) {
                const cleanName = newBrandName.trim();

                try {
                    // Create the brand in the lookup list
                    await createBrandInList(cleanName);

                    // Reload and repopulate the brand select
                    const allBrands = await loadAllBrandsList();
                    cache.brandsList = allBrands;
                    populateBrandSelect(select, allBrands);

                    // Set the newly added brand as selected
                    select.value = cleanName;
                    value = cleanName;
                    text = cleanName;
                } catch (err) {
                    console.error("Failed to add new brand to list", err);
                    alert("Failed to add new brand. Please try again.");
                    select.value = "";
                    return;
                }
            } else {
                select.value = "";
                return;
            }
        }

        if (!value || !text || value === "__add_new__") return;

        const already = [...wrap.querySelectorAll(".brand-chip")]
            .some(ch => (ch.dataset.value || "").toLowerCase() === text.toLowerCase());
        if (already) {
            select.value = "";
            return;
        }

        wrap.appendChild(buildBrandChip(text));
        select.value = "";
    }

    function appendOrderRow(orderNumber, labelText, custPoNumber = "") {
        const tbody = $("#orderNumbersTable tbody");
        const tableSection = $("#orderNumbersTableSection");
        if (!tbody || !tableSection) return;

        if ([...tbody.querySelectorAll("tr")].some(tr => tr.dataset.order === orderNumber)) {
            return;
        }

        const tr = d.createElement("tr");
        tr.dataset.order = orderNumber;
        tr.dataset.ciNumber = custPoNumber || "";  // ← STORE CI# IN DATASET

        const tdOrder = d.createElement("td");
        tdOrder.style.padding = "10px";
        tdOrder.textContent = labelText || orderNumber;

        const tdCI = d.createElement("td");
        tdCI.style.padding = "10px";
        tdCI.textContent = custPoNumber || "";

        const tdAct = d.createElement("td");
        tdAct.style.padding = "10px";

        const rm = d.createElement("button");
        rm.type = "button";
        rm.textContent = "Remove";
        rm.className = "btn btn-outline";
        rm.addEventListener("click", () => {
            tr.remove();
            if (!tbody.children.length) {
                tableSection.style.display = "none";
            }
        });

        tdAct.appendChild(rm);

        tr.appendChild(tdOrder);
        tr.appendChild(tdCI);
        tr.appendChild(tdAct);

        tbody.appendChild(tr);
        tableSection.style.display = "";
    }

    function addOrderFromDropdown() {
        const sel = $("#orderNumberSelect");
        if (!sel) return;
        const orderNumber = sel.value;
        const selectedOption = sel.options[sel.selectedIndex];
        const label = selectedOption?.textContent?.trim() || "";
        const custPoNumber = selectedOption?.dataset.custPoNumber || "";  // ← GET custPoNumber
        if (!orderNumber) return;

        appendOrderRow(orderNumber, label, custPoNumber);  // ← PASS custPoNumber
        sel.value = "";
    }

    /************************************************************
     * 9. COLLECT DATA FOR CREATE
     ************************************************************/
    function collectPartyBlock() {
        const partyType = ($("#shipToLabel")?.value || "shipto").toLowerCase();
        const mainText = $("#shipTo")?.value || "";
        const billToVal = $("#billTo")?.value || "";

        const out = {
            partyLabel: "",
            shiptolocation: "",
            billto: "",
            consignee: "",
            applicant: ""
        };

        if (partyType === "consignee") {
            out.partyLabel = "Consignee";
            out.consignee = mainText;
        } else if (partyType === "applicant") {
            out.partyLabel = "Applicant";
            out.applicant = mainText;
        } else {
            out.partyLabel = "Ship To";
            out.shiptolocation = mainText;
            out.billto = billToVal;
        }

        return out;
    }

    function collectHeaderPayload() {
        // Party fields removed - now using Customer Models instead

        // Get country and generate country code
        const countryValue = $("#country")?.value || "";
        const countryCode = getCountryCode(countryValue);

        // Set initial DCL number with country code (will be completed after creation)
        const initialDclNumber = `DCL-${countryCode}-`;

        const body = {
            cr650_dclnumber: initialDclNumber,  // ← ADD THIS LINE
            cr650_customername: cache.selectedCustomer?.name || "",
            cr650_customernumber: $("#customerNumber")?.value || "",
            cr650_company_type: $("#companyType")?.value ? parseInt($("#companyType").value) : null,

            cr650_country: $("#country")?.value || "",
            cr650_organizationid: $("#orgID")?.value || "",
            cr650_cob: $("#COB")?.value || "",
            cr650_countryoforigin: $("#COO")?.value || "",
            cr650_paymentterms: $("#PaymentTerms")?.value || "",
            cr650_currencycode: $("#currency")?.value || "",
            cr650_incoterms: $("#incoterms")?.value || "",

            cr650_shippingline: $("#shippingLineSelect")?.value || "",
            cr650_loadingport: $("#loadingPort")?.value || "",
            cr650_destinationport: $("#destinationPort")?.value || "",
            cr650_actualreadinessdate: asISO($("#actualRednessDate")?.value),
            cr650_loadingdate: asISO($("#loadingDate")?.value),
            cr650_sailing_date: asISO($("#sailingDate")?.value),

            cr650_loadingtime: $("#loadingTime")?.value || "",
            cr650_loading_shift: $("#loadingShift")?.value || "",
            cr650_transportationmode: $("#transportation")?.value || "",
            cr650_salesrepresentativename: $("#salesRep")?.value || "",
            cr650_description_goods_services: $("#descriptionGoodsServices")?.value || ""

            // Party fields removed - now using Customer Models instead
        };

        Object.keys(body).forEach(k => {
            if (
                body[k] === "" ||
                body[k] === undefined ||
                body[k] === null
            ) {
                delete body[k];
            }
        });

        return body;
    }

    function collectNotifyPartiesFromUI() {
        const out = [];
        $$("#notifyPartyContainer .notify-card").forEach(card => {
            const ta = card.querySelector(".notify-text");
            const txt = (ta?.value || "").trim();
            if (!txt) return;
            out.push({
                cr650_notify_party: txt
            });
        });
        return out;
    }

    function collectTermsFromUI() {
        const out = [];
        $$("#termsConditionsContainer .terms-condition-item").forEach(item => {
            const ta = item.querySelector(".terms-textarea");
            const cbCI = item.querySelector(".tc-print-ci");
            const cbPL = item.querySelector(".tc-print-pl");

            const textVal = (ta?.value || "").trim();
            if (!textVal) return;

            out.push({
                cr650_termcondition: textVal,
                cr650_is_printed_ci: !!cbCI?.checked,
                cr650_is_printed_pl: !!cbPL?.checked
            });
        });
        return out;
    }

    function collectBrandsFromUI() {
        const out = [];
        $$("#brandList .brand-chip").forEach(chip => {
            const brandName = chip.dataset.value || chip.textContent || "";
            const clean = brandName.trim();
            if (!clean) return;
            out.push({
                cr650_brand: clean
            });
        });
        return out;
    }

    function collectOrdersFromUI() {
        const out = [];
        $$("#orderNumbersTable tbody tr").forEach(tr => {
            const orderNum = tr.dataset.order || "";
            const ciNum = tr.dataset.ciNumber || "";  // ← GET CI# FROM DATASET
            if (!orderNum) return;
            out.push({
                cr650_order_number: orderNum,
                cr650_ci_number: ciNum  // ← INCLUDE CI# IN COLLECTION
            });
        });
        return out;
    }







    /************************************************************
     * 10. CREATION PIPELINE
     ************************************************************/
    async function createDclMaster(headerPayload) {
        async function doRequest(token) {
            const headers = {
                "Accept": "application/json;odata.metadata=minimal",
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "return=representation"
            };
            if (token) {
                headers["__RequestVerificationToken"] = token;
            }

            TopLoader.begin();
            try {
                // Step 1: Create the initial record
                const res = await fetch(DCL_API, {
                    method: "POST",
                    headers,
                    body: JSON.stringify(headerPayload)
                });

                const rawText = await res.text();
                let jsonBody = {};
                try {
                    jsonBody = rawText ? JSON.parse(rawText) : {};
                } catch {
                    jsonBody = {};
                }

                if (!res.ok) {
                    throw new Error(rawText || ("HTTP " + res.status));
                }

                // Extract the new record ID
                let newId =
                    jsonBody.cr650_dcl_masterid ||
                    jsonBody.cr650_dcl_masterId ||
                    jsonBody.cr650_dclmasterid ||
                    jsonBody.cr650_dclmasterId ||
                    jsonBody.id ||
                    "";

                if ((!newId || newId === "") && jsonBody["@odata.id"]) {
                    const m = jsonBody["@odata.id"].match(/\(([^)]+)\)$/);
                    if (m && m[1]) {
                        newId = m[1];
                    }
                }

                if (!newId || newId === "") {
                    const entityIdHeader =
                        res.headers.get("OData-EntityId") ||
                        res.headers.get("odata-entityid") ||
                        res.headers.get("odata-entity-id") ||
                        res.headers.get("OData-EntityID");
                    if (entityIdHeader) {
                        const m2 = entityIdHeader.match(/\(([^)]+)\)$/);
                        if (m2 && m2[1]) {
                            newId = m2[1];
                        }
                    }
                }

                if (typeof newId === "string") {
                    newId = newId.replace(/[{}]/g, "");
                }

                if (!newId) {
                    throw new Error("Failed to extract record ID from creation response");
                }

                // Step 2: Fetch the record to get the auto-number
                const fetchUrl = `${DCL_API}(${newId})?$select=cr650_autonumber,cr650_dclnumber`;
                const fetchRes = await fetch(fetchUrl, {
                    method: "GET",
                    headers: {
                        "Accept": "application/json;odata.metadata=minimal"
                    }
                });

                if (!fetchRes.ok) {
                    console.warn("Failed to fetch auto-number, proceeding with partial DCL number");
                    return {
                        newId,
                        dclNumber: headerPayload.cr650_dclnumber || "Pending"
                    };
                }

                const fetchData = await fetchRes.json();
                const autoNumber = fetchData.cr650_autonumber;

                if (!autoNumber) {
                    console.warn("Auto-number not yet generated, proceeding with partial DCL number");
                    return {
                        newId,
                        dclNumber: headerPayload.cr650_dclnumber || "Pending"
                    };
                }

                // Step 3: Format the auto-number with leading zeros
                const formattedAutoNumber = String(autoNumber).padStart(4, '0');

                // Step 4: Construct the complete DCL number
                const completeDclNumber = `${headerPayload.cr650_dclnumber}${formattedAutoNumber}`;

                // Step 5: Update the record with the complete DCL number
                const updateHeaders = {
                    "Accept": "application/json;odata.metadata=minimal",
                    "Content-Type": "application/json; charset=utf-8",
                    "If-Match": "*"
                };
                if (token) {
                    updateHeaders["__RequestVerificationToken"] = token;
                }

                const updateRes = await fetch(`${DCL_API}(${newId})`, {
                    method: "PATCH",
                    headers: updateHeaders,
                    body: JSON.stringify({
                        cr650_dclnumber: completeDclNumber
                    })
                });

                if (!updateRes.ok) {
                    console.warn("Failed to update DCL number with auto-number");
                }

                return {
                    newId,
                    dclNumber: completeDclNumber
                };

            } finally {
                TopLoader.end();
            }
        }

        if (w.shell && typeof shell.getTokenDeferred === "function") {
            return new Promise((resolve, reject) => {
                shell.getTokenDeferred()
                    .done((token) => {
                        doRequest(token).then(resolve).catch(reject);
                    })
                    .fail(() => {
                        reject(new Error("Token API unavailable (user not authenticated?)"));
                    });
            });
        } else {
            return doRequest(null);
        }
    }

    async function createNotifyPartyRows(dclId, list) {
        for (const row of list) {
            const bodyObj = Object.assign({}, row, {
                "cr650_Dcl_number@odata.bind": "/cr650_dcl_masters(" + dclId + ")"
            });

            await safeAjax({
                type: "POST",
                url: NOTIFY_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });
        }
    }

    async function createTermsRows(dclId, list) {
        for (const row of list) {
            const bodyObj = {
                cr650_termcondition: row.cr650_termcondition,
                cr650_is_printed_ci: !!row.cr650_is_printed_ci,
                cr650_is_printed_pl: !!row.cr650_is_printed_pl,
                "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + dclId + ")"
            };

            await safeAjax({
                type: "POST",
                url: TERMS_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });
        }
    }

    async function createBrandRows(dclId, list) {
        for (const row of list) {
            const bodyObj = {
                cr650_brand: row.cr650_brand,
                "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + dclId + ")"
            };

            await safeAjax({
                type: "POST",
                url: BRANDS_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });
        }
    }

    async function createOrderRows(dclId, list) {
        for (const row of list) {
            const bodyObj = {
                cr650_order_number: row.cr650_order_number,
                cr650_ci_number: row.cr650_ci_number || "",  // ← ADD CI# FIELD
                "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + dclId + ")"
            };

            await safeAjax({
                type: "POST",
                url: DCL_ORDERS_API,
                headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    "Prefer": "return=representation"
                },
                data: JSON.stringify(bodyObj)
            });
        }
    }

    /************************************************************
 * DOCUMENT NUMBER EXTRACTION VIA POWER AUTOMATE
 ************************************************************/
    const POWER_AUTOMATE_ENDPOINT = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a30fd12c4ca34488b589d3261b5a2eba/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=X3vQ-GwyonKHmo-ZImPVA45_eYWelJgKYQU_tR6Vm2Y";

    /**
     * Convert file to base64
     */
    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                // Remove the data URL prefix (e.g., "data:application/pdf;base64,")
                const base64 = reader.result.split(',')[1];
                resolve(base64);
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    /**
     * Get file extension from filename
     */
    function getFileExtension(filename) {
        const parts = filename.split('.');
        return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
    }

    /**
     * Send file to Power Automate for number extraction
     */
    async function extractNumberFromFile(file) {
        try {
            TopLoader.begin();

            // Convert file to base64
            const base64 = await fileToBase64(file);
            const extension = getFileExtension(file.name);

            // Prepare payload
            const payload = {
                file_base64: base64,
                file_extension: extension
            };

            // Send to Power Automate
            const response = await fetch(POWER_AUTOMATE_ENDPOINT, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const data = await response.json();
            return data;

        } catch (error) {
            console.error('Failed to extract number from file:', error);
            throw error;
        } finally {
            TopLoader.end();
        }
    }

   /**
 * Create a dropdown with extracted numbers
 * FIXED: Now handles both string values and object values
 */
function createNumberDropdown(numbers, targetInputId) {
    const targetInput = document.getElementById(targetInputId);
    if (!targetInput) {
        console.error(`Target input ${targetInputId} not found`);
        return;
    }

    // Remove existing dropdown if any
    const existingDropdown = document.getElementById(`${targetInputId}_dropdown`);
    if (existingDropdown) {
        existingDropdown.remove();
    }

    // Create select element
    const select = document.createElement('select');
    select.id = `${targetInputId}_dropdown`;
    select.style.cssText = `
        padding: 0.875rem 1rem;
        border-radius: 12px;
        font-size: 0.95rem;
        margin-top: 8px;
        width: 100%;
        border: 1px solid #d1d5db;
    `;

    // Add default option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select extracted number...';
    select.appendChild(defaultOption);

    // Add options for each extracted number
    // FIXED: Handle both string and object formats
    numbers.forEach((item, index) => {
        const option = document.createElement('option');
        
        // Handle both string and object formats
        const numberValue = typeof item === 'string' 
            ? item 
            : (item.item || item.number || item.value || '');
        
        option.value = numberValue;
        option.textContent = numberValue || `Unknown Item ${index + 1}`;
        select.appendChild(option);

        console.log(`Added option: ${numberValue}`); // Debug log
    });

    // Handle selection
    select.addEventListener('change', function () {
        const selectedValue = this.value;
        console.log(`Selected value: ${selectedValue}`); // Debug log

        if (selectedValue) {
            targetInput.value = selectedValue;

            // Trigger a change event on the input
            const changeEvent = new Event('change', { bubbles: true });
            targetInput.dispatchEvent(changeEvent);

            // Trigger autosave if DCL exists
            if (window.currentDclId) {
                const fieldMap = {
                    'edNumber': 'cr650_ednumber',
                    'blNumber': 'cr650_blnumber',
                    'piNumber': 'cr650_pinumber',
                    'poNumber': 'cr650_ponumber'
                };

                const dataverseField = fieldMap[targetInputId];
                if (dataverseField) {
                    console.log(`Saving to ${dataverseField}: ${selectedValue}`); // Debug log
                    patchDclField(window.currentDclId, {
                        [dataverseField]: selectedValue
                    });
                }
            }
        }
    });

    // Insert after the file input (not the text input)
    const fileInputId = targetInputId.replace('Number', 'File');
    const fileInput = document.getElementById(fileInputId);

    if (fileInput && fileInput.parentNode) {
        fileInput.parentNode.insertBefore(select, fileInput.nextSibling);
    } else {
        // Fallback: insert after text input
        targetInput.parentNode.insertBefore(select, targetInput.nextSibling);
    }
}


    /**
     * Handle file upload and number extraction
     */
 
async function handleDocumentFileUpload(fileInput, textInputId) {
    const file = fileInput.files[0];
    if (!file) return;

    const textInput = document.getElementById(textInputId);
    if (!textInput) {
        console.error(`Text input ${textInputId} not found`);
        return;
    }

    try {
        // Show loading state
        const originalValue = textInput.value;
        textInput.value = 'Extracting number...';
        textInput.disabled = true;

        // Extract numbers
        const response = await extractNumberFromFile(file);

        console.log('Extraction response:', response); // Debug log

        // Handle response
        if (response && response.numbers && Array.isArray(response.numbers)) {
            const numbers = response.numbers;

            console.log(`Found ${numbers.length} numbers:`, numbers); // Debug log

            if (numbers.length === 0) {
                // No numbers found
                alert('No numbers were extracted from the document. Please enter manually.');
                textInput.value = originalValue;
            } else if (numbers.length === 1) {
                // Single number - populate directly
                // FIXED: Handle both string and object formats
                const firstItem = numbers[0];
                const extractedNumber = typeof firstItem === 'string' 
                    ? firstItem 
                    : (firstItem.item || firstItem.number || firstItem.value || '');
                
                textInput.value = extractedNumber;

                console.log(`✅ Setting single value: ${extractedNumber}`); // Debug log

                // Trigger autosave if DCL exists
                if (window.currentDclId) {
                    const fieldMap = {
                        'edNumber': 'cr650_ednumber',
                        'blNumber': 'cr650_blnumber',
                        'piNumber': 'cr650_pinumber',
                        'poNumber': 'cr650_ponumber'
                    };

                    const dataverseField = fieldMap[textInputId];
                    if (dataverseField) {
                        console.log(`💾 Auto-saving to ${dataverseField}: ${extractedNumber}`);
                        await patchDclField(window.currentDclId, {
                            [dataverseField]: extractedNumber
                        });
                    }
                }
            } else {
                // Multiple numbers - show dropdown
                console.log(`📋 Creating dropdown with ${numbers.length} options`);
                textInput.value = originalValue; // Restore original value
                createNumberDropdown(numbers, textInputId);
            }
        } else {
            throw new Error('Invalid response format: ' + JSON.stringify(response));
        }

    } catch (error) {
        console.error('❌ Error processing document:', error);
        alert('Failed to extract number from document. Please enter manually.\n\nError: ' + error.message);
        textInput.value = '';
    } finally {
        textInput.disabled = false;
    }
}

    /**
     * Wire up file upload handlers for document number extraction
     */
    function initDocumentNumberExtraction() {
    }

    // Expose function globally
    w.handleDocumentFileUpload = handleDocumentFileUpload;
    // Expose function globally


    /************************************************************
     * 11. CREATE BUTTON HANDLER
     ************************************************************/
    async function onCreateNewDclClick() {
        // use helper instead of #customerNameText (which may not exist)
        const custNameVal = getSelectedCustomerNameOrEmpty().trim();
        const incotermVal = ($("#incoterms")?.value || "").trim();

        if (!custNameVal) {
            alert("Please select a Customer.");
            return;
        }
        if (!incotermVal) {
            alert("Incoterms is required.");
            return;
        }

        const headerPayload = collectHeaderPayload();

        // Inject submitter info before sending to Dataverse
        const emailField = document.getElementById("user-emailaddress");
        const usernameField = document.getElementById("user-username");

        if (emailField && emailField.value.trim()) {
            headerPayload.cr650_submitter_email = emailField.value.trim();
        }
        if (usernameField && usernameField.value.trim()) {
            headerPayload.cr650_submitter_name = usernameField.value.trim();
        }


        const ordersList = collectOrdersFromUI();
        const notifyList = collectNotifyPartiesFromUI();
        const termsList = collectTermsFromUI();
        const brandsList = collectBrandsFromUI();

        setBusy(true, "Creating DCL…");

        try {
            const { newId, dclNumber } = await createDclMaster(headerPayload);

            w.currentDclId = newId;

            if ($("#dclNumber")) {
                $("#dclNumber").textContent = dclNumber || "(Pending)";
            }

            if (!newId) {
                setBusy(false);
                alert("DCL created, but ID not returned. Please open it from the list.");
                return;
            }

            if (notifyList.length) {
                await createNotifyPartyRows(newId, notifyList);
            }
            if (termsList.length) {
                await createTermsRows(newId, termsList);
            }
            if (brandsList.length) {
                await createBrandRows(newId, brandsList);
            }
            if (ordersList.length) {
                await createOrderRows(newId, ordersList);
            }

            // Create pending customer models
            if (cache.pendingDCLModels.length > 0) {
                console.log("Creating pending DCL models:", cache.pendingDCLModels.length);
                for (const pendingModel of cache.pendingDCLModels) {
                    try {
                        await createCustomerModelForDCL(newId, pendingModel.modelName, pendingModel.info);
                        console.log("Created model:", pendingModel.modelName);
                    } catch (modelErr) {
                        console.error("Failed to create model:", pendingModel.modelName, modelErr);
                    }
                }
                cache.pendingDCLModels = [];
            }

            TopLoader.begin();
            window.location.href =
                "/Basic_Information_View/?id=" +
                encodeURIComponent(newId);

        } catch (err) {
            console.error("Failed to create new DCL", err);
            alert("Error creating DCL. Check required fields and permissions.");
            setBusy(false);
            TopLoader.forceClear();
        }
    }

    /************************************************************
     * 12. AUTOSAVE FIELD BINDINGS
     ************************************************************/
    const FIELD_BINDINGS = [
        { selector: "#shippingLineSelect", attr: "cr650_shippingline", event: "change" },
        { selector: "#loadingPort", attr: "cr650_loadingport", event: "blur" },
        { selector: "#destinationPort", attr: "cr650_destinationport", event: "blur" },
        { selector: "#loadingDate", attr: "cr650_loadingdate", event: "change", transform: asISO },
        { selector: "#loadingTime", attr: "cr650_loadingtime", event: "change" },
        { selector: "#loadingShift", attr: "cr650_loading_shift", event: "change" },
        { selector: "#transportation", attr: "cr650_transportationmode", event: "change" },

        { selector: "#country", attr: "cr650_country", event: "blur" },
        { selector: "#orgID", attr: "cr650_organizationid", event: "blur" },
        { selector: "#COB", attr: "cr650_cob", event: "blur" },
        { selector: "#COO", attr: "cr650_countryoforigin", event: "blur" },
        { selector: "#PaymentTerms", attr: "cr650_paymentterms", event: "blur" },
        { selector: "#currency", attr: "cr650_currencycode", event: "change" },
        { selector: "#incoterms", attr: "cr650_incoterms", event: "change" },
        { selector: "#salesRep", attr: "cr650_salesrepresentativename", event: "blur" },
        {
            selector: "#companyType",
            attr: "cr650_company_type",
            event: "change",
            transform: (val) => val ? parseInt(val) : null
        },
        // Party fields removed - now using Customer Models instead
        {
            selector: "#customerNumber", attr: "cr650_customernumber", event: "blur",
            afterSave: async (val) => { await reloadOrdersForNewCustomerNumber(val); }
        }
    ];

    function wireAutoSaveFields() {
        FIELD_BINDINGS.forEach(binding => {
            const fieldEl = $(binding.selector);
            if (!fieldEl) return;

            const handler = async () => {
                const rawVal = fieldEl.value;
                const val = binding.transform ? binding.transform(rawVal) : rawVal;

                const body = {};
                if (val !== "" && val !== undefined && val !== null) {
                    body[binding.attr] = val;
                } else {
                    body[binding.attr] = null;
                }

                await patchDclField(w.currentDclId, body);

                if (binding.afterSave) {
                    try {
                        await binding.afterSave(rawVal);
                    } catch (e) {
                        console.error("afterSave handler error:", e);
                    }
                }
            };

            fieldEl.addEventListener(binding.event, handler);
        });

        // hook loadingTimeType here for UI only
        const loadingTypeSel = $("#loadingTimeType");
        if (loadingTypeSel) {
            loadingTypeSel.addEventListener("change", toggleLoadingTimeInput);
        }
    }

    /************************************************************
     * 13. SPECIAL HANDLERS
     ************************************************************/
    async function handleCustomerChange() {
        const selNode = el.customerPicker;
        const idAttr = selNode?.value || "";

        if (!idAttr) {
            cache.selectedCustomer = null;
            cache.customerModels = [];

            setVal("#customerNameText", "");
            setVal("#customerNumber", "");
            setVal("#country", "");
            setVal("#orgID", "");
            setVal("#orderType", "");
            setVal("#COB", "");
            setVal("#PaymentTerms", "");
            setVal("#salesRep", "");
            setVal("#destinationPort", "");
            setVal("#currency", "");
            clearOrderSelectAndTable();

            // Clear customer models section
            renderCustomerModelsSection([]);

            await patchDclField(w.currentDclId, {
                cr650_customername: null,
                cr650_customernumber: null
            });

            return;
        }

        const idx = selNode.options[selNode.selectedIndex]?.dataset.idx;
        const cust = cache.customers[idx]
            || cache.customers.find(c => c.id === idAttr)
            || null;

        cache.selectedCustomer = cust;

        if (cust) {
            trySelectCustomerOnUI(cust);

            tryDefaultCurrencyByCountry();

            // Auto-populate notify parties from customer defaults
            populateNotifyPartiesFromCustomer(cust);

            // Load customer models for this customer
            try {
                const models = await loadCustomerModelsForCustomer(cust.id);
                cache.customerModels = models;
                renderCustomerModelsSection(models);
            } catch (err) {
                console.error("Failed to load customer models:", err);
                renderCustomerModelsSection([]);
            }
        }

        await patchDclField(w.currentDclId, {
            cr650_customername: cust?.name || "",
            cr650_customernumber: cust?.code || ""
        });

        const custNumber = (cust?.code || "").trim();
        if (!custNumber) {
            clearOrderSelectAndTable();
        } else {
            await reloadOrdersForNewCustomerNumber(custNumber);
        }
    }

    /**
     * Populate notify parties from customer's default notify party fields
     */
    function populateNotifyPartiesFromCustomer(cust) {
        const container = $("#notifyPartyContainer");
        if (!container) return;

        const np1 = (cust.notifyParty1 || "").trim();
        const np2 = (cust.notifyParty2 || "").trim();

        // If customer has no notify parties, keep existing (don't clear)
        if (!np1 && !np2) {
            return;
        }

        // Clear existing notify party rows
        container.innerHTML = "";

        // Add notify party 1 if exists
        if (np1) {
            const row1 = buildNotifyPartyRow(np1);
            container.appendChild(row1);
        }

        // Add notify party 2 if exists
        if (np2) {
            const row2 = buildNotifyPartyRow(np2);
            container.appendChild(row2);
        }
    }

    function onShipToLabelChange() {
        const mode = ($("#shipToLabel")?.value || "").toLowerCase();

        const src = cache.selectedCustomer || {};
        fillPartyUIFromRecord(mode, {
            shipTo: src.shipTo,
            billTo: src.billTo,
            consignee: src.consignee,
            primaryAddress: src.primaryAddress
        });

        const mainText = $("#shipTo")?.value || "";
        const billToVal = $("#billTo")?.value || "";

        let body = {
            cr650_party: "",
            cr650_shiptolocation: null,
            cr650_billto: null,
            cr650_consignee: null,
            cr650_applicant: null
        };

        if (mode === "consignee") {
            body.cr650_party = "Consignee";
            body.cr650_consignee = mainText || "";
        } else if (mode === "applicant") {
            body.cr650_party = "Applicant";
            body.cr650_applicant = mainText || "";
        } else {
            body.cr650_party = "Ship To";
            body.cr650_shiptolocation = mainText || "";
            body.cr650_billto = billToVal || "";
        }

        patchDclField(w.currentDclId, body);
    }

    function toggleLoadingTimeInput() {
        const typeSel = $("#loadingTimeType");
        const mode = typeSel ? (typeSel.value || "") : "";

        const specific = $("#specificTimeInput");
        const shift = $("#shiftInput");

        if (specific) {
            specific.style.display = (mode === "specific") ? "" : "none";
        }
        if (shift) {
            shift.style.display = (mode === "shift") ? "" : "none";
        }
    }
    w.toggleLoadingTimeInput = toggleLoadingTimeInput;

    async function handleShippingLineChange(selectEl) {
        if (!selectEl) return;

        if (selectEl.value === "__add_new__") {
            const name = prompt("Enter new shipping line name:");

            if (name && name.trim()) {
                const clean = name.trim();

                try {
                    // Save to database
                    TopLoader.begin();
                    const newLine = await createShippingLine(clean);

                    // Reload all shipping lines to refresh the dropdown
                    const allLines = await loadAllShippingLines();
                    cache.shippingLines = allLines;
                    populateShippingLineSelect(selectEl, allLines);

                    // Select the newly added line
                    selectEl.value = newLine.name;

                    // Trigger autosave if DCL exists (for edit mode)
                    if (w.currentDclId) {
                        await patchDclField(w.currentDclId, {
                            cr650_shippingline: newLine.name
                        });
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
    }
    w.handleShippingLineChange = handleShippingLineChange;

    /************************************************************
     * 14. GLOBAL CLICK for remove buttons
     ************************************************************/
    d.addEventListener("click", (e) => {
        if (e.target.classList.contains("notify-delete-btn")) {
            const card = e.target.closest(".notify-card");
            if (!card) return;
            card.remove();
            return;
        }

        if (e.target.classList.contains("tc-delete-btn")) {
            const card = e.target.closest(".terms-condition-item");
            if (!card) return;
            card.remove();
            return;
        }

        if (e.target.classList.contains("brand-remove-btn")) {
            const chip = e.target.closest(".brand-chip");
            if (!chip) return;
            chip.remove();
            return;
        }
    });

    /************************************************************
     * 15. COLLAPSIBLES
     ************************************************************/
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
            if (ic) ic.style.transform = isOpen ? "rotate(0deg)" : "rotate(-90deg)";
        });

        d.addEventListener("click", (e) => {
            const header = e.target.closest(".section-header");
            if (!header) return;

            const card = header.closest(".info-section");
            const content = card?.querySelector(".section-content");
            if (!content) return;

            const willOpen = !(card && card.classList.contains("open"));
            card.classList.toggle("open", willOpen);

            if (willOpen) {
                content.style.display = "block";
                header.setAttribute("aria-expanded", "true");
            } else {
                content.style.display = "none";
                header.setAttribute("aria-expanded", "false");
            }

            const icon = header.querySelector(".toggle-icon");
            if (icon) {
                icon.style.transform = willOpen ? "rotate(0deg)" : "rotate(-90deg)";
            }
        });
    }

    /************************************************************
     * 16. BOOTSTRAP
     ************************************************************/
    /************************************************************
 * 16. BOOTSTRAP
 ************************************************************/
    d.addEventListener("DOMContentLoaded", async () => {
        // ============================================================
        // 1. CACHE DOM REFS
        // ============================================================
        el.customerPicker = $("#customerPicker");
        el.companyType = $("#companyType");

        // ============================================================
        // 2. WIRE CUSTOMER PICKER
        // ============================================================
        if (el.customerPicker) {
            el.customerPicker.addEventListener("change", handleCustomerChange);
        }

        // ============================================================
        // 3. WIRE COMPANY TYPE (with color change & auto-save)
        // ============================================================
        if (el.companyType) {
            el.companyType.addEventListener("change", () => {
                updateCompanyTypeColor();

                // Auto-save if DCL exists
                if (w.currentDclId) {
                    const val = el.companyType.value;
                    patchDclField(w.currentDclId, {
                        cr650_company_type: val ? parseInt(val) : null
                    });
                }
            });

            // Initialize color on load
            updateCompanyTypeColor();
        }

        // ============================================================
        // 4. WIRE CUSTOMER MODELS
        // ============================================================
        const addModelBtn = $("#btnAddCustomerModel");
        if (addModelBtn) {
            addModelBtn.addEventListener("click", handleAddModelToDCL);
        }

        // Event delegation for "Use" buttons on customer models
        d.addEventListener("click", async function(e) {
            // Use model button
            if (e.target.classList.contains("btn-use-model") || e.target.closest(".btn-use-model")) {
                const btn = e.target.classList.contains("btn-use-model") ? e.target : e.target.closest(".btn-use-model");
                const modelName = btn.getAttribute("data-model-name") || "";
                const modelInfo = btn.getAttribute("data-model-info") || "";

                // Pre-fill the add form with this model's data
                const nameInput = $("#newModelName");
                const infoInput = $("#newModelInfo");
                if (nameInput) nameInput.value = modelName;
                if (infoInput) infoInput.value = modelInfo;

                // Scroll to the form
                const addForm = $("#addModelForm");
                if (addForm) addForm.scrollIntoView({ behavior: "smooth", block: "center" });
            }

            // Delete DCL model button
            if (e.target.classList.contains("btn-delete-dcl-model") || e.target.closest(".btn-delete-dcl-model")) {
                const btn = e.target.classList.contains("btn-delete-dcl-model") ? e.target : e.target.closest(".btn-delete-dcl-model");
                const modelId = btn.getAttribute("data-model-id");
                if (!modelId) return;

                if (!confirm("Remove this model?")) return;

                // Handle pending models (not yet saved to DB)
                if (modelId.startsWith('pending-')) {
                    cache.pendingDCLModels = cache.pendingDCLModels.filter(m => m.id !== modelId);
                    renderDCLModelsSection(cache.pendingDCLModels);
                    return;
                }

                // Handle saved models
                try {
                    await deleteCustomerModelFromDCL(modelId);
                    await refreshDCLModels();
                } catch (err) {
                    console.error("Failed to delete model:", err);
                    alert("Failed to delete model. Please try again.");
                }
            }
        });

        // ============================================================
        // 5. INITIALIZE LOADING TIME VISIBILITY
        // ============================================================
        toggleLoadingTimeInput();

        // ============================================================
        // 6. WIRE ACTION BUTTONS (Add Notify, Terms, Brand, Order)
        // ============================================================
        const addNotifyBtn = $("#btnAddNotifyParty");
        if (addNotifyBtn) {
            addNotifyBtn.addEventListener("click", appendNotifyPartyRow);
        }

        const addTermsBtn = $("#addTermsConditionBtn");
        if (addTermsBtn) {
            addTermsBtn.addEventListener("click", appendTermsConditionRow);
        }

        const addBrandBtn = $("#addBrandBtn");
        if (addBrandBtn) {
            addBrandBtn.addEventListener("click", addBrandChipFromSelect);
        }

        const addOrderBtn = $("#addOrderBtn");
        if (addOrderBtn) {
            addOrderBtn.addEventListener("click", addOrderFromDropdown);
        }

        // ============================================================
        // 7. WIRE CREATE DCL BUTTON
        // ============================================================
        const createBtn = d.querySelector(".navigation-buttons .btn.btn-primary");
        if (createBtn) {
            createBtn.addEventListener("click", onCreateNewDclClick);
        }

        // ============================================================
        // 8. INITIALIZE COLLAPSIBLE SECTIONS
        // ============================================================
        initCollapsibles();

        // ============================================================
        // 9. ADD DEFAULT EMPTY ROWS (Notify Party & Terms)
        // ============================================================
        appendNotifyPartyRow();
        appendTermsConditionRow();

        // ============================================================
        // 10. LOAD CURRENCIES (with fallback)
        // ============================================================
        try {
            await ensureCurrenciesLoaded();
        } catch (errCurrency) {
            console.warn("Currency API failed, using fallback list:", errCurrency);
            cache.currencies = [
                { code: "USD", name: "United States Dollar" },
                { code: "SAR", name: "Saudi Riyal" },
                { code: "EUR", name: "Euro" },
                { code: "AED", name: "UAE Dirham" },
                { code: "GBP", name: "Pound Sterling" }
            ];
        }

        // Populate all currency dropdowns
        populateAllCurrencySelects();
        // ============================================================
        // 10.5. LOAD SHIPPING LINES
        // ============================================================
        try {
            const shippingSelect = $("#shippingLineSelect");
            if (shippingSelect) {
                console.log("🔍 Starting shipping lines load...");

                shippingSelect.innerHTML = '<option value="">Loading shipping lines…</option>';
                shippingSelect.disabled = true;

                const shippingLines = await loadAllShippingLines();
                cache.shippingLines = shippingLines;

                console.log("✅ Loaded shipping lines:", shippingLines.length);

                populateShippingLineSelect(shippingSelect, shippingLines);
                shippingSelect.disabled = false;
            }
        } catch (errShipping) {
            console.error("❌ Failed to load shipping lines:", errShipping);
            const shippingSelect = $("#shippingLineSelect");
            if (shippingSelect) {
                shippingSelect.innerHTML = '<option value="">Failed to load shipping lines</option>';
                shippingSelect.disabled = false;
            }
        }

        // ============================================================
        // 10.6. LOAD BRANDS LIST
        // ============================================================
        try {
            const brandSelect = $("#brandSelect");
            if (brandSelect) {
                console.log("🔍 Starting brands list load...");
                brandSelect.innerHTML = '<option value="">Loading brands…</option>';
                brandSelect.disabled = true;

                const brandsList = await loadAllBrandsList();
                cache.brandsList = brandsList;
                populateBrandSelect(brandSelect, brandsList);

                brandSelect.disabled = false;
                console.log("✅ Loaded brands:", brandsList.length);
            }
        } catch (errBrands) {
            console.error("❌ Failed to load brands:", errBrands);
            const brandSelect = $("#brandSelect");
            if (brandSelect) {
                brandSelect.innerHTML = '<option value="">Failed to load brands</option>';
                brandSelect.disabled = false;
            }
        }

        // ============================================================
        // 11. LOAD CUSTOMERS
        // ============================================================
        try {
            if (el.customerPicker) {
                console.log("🔍 Starting customer load...");

                el.customerPicker.disabled = true;
                el.customerPicker.innerHTML = '<option value="">Loading customers…</option>';

                console.log("🔍 Calling loadAllCustomers()...");
                const list = await loadAllCustomers();

                console.log("✅ Loaded customers:", list.length);
                console.log("✅ First 3 customers:", list.slice(0, 3));

                populateCustomerSelect(list);

                console.log("✅ Picker populated with", el.customerPicker.options.length, "options");
            } else {
                console.warn("⚠️ customerPicker element not found!");
            }
        } catch (errCust) {
            console.error("❌ Failed to load customers:", errCust);
            console.error("❌ Error details:", errCust.message, errCust.stack);

            if (el.customerPicker) {
                el.customerPicker.innerHTML = '<option value="">Failed to load customers</option>';
                el.customerPicker.disabled = true;
            }
        }

        // ============================================================
        // 12. WIRE AUTOSAVE FIELD BINDINGS
        // ============================================================
        wireAutoSaveFields();

        initDocumentNumberExtraction();


        // ============================================================
        // 13. CLEANUP: REMOVE BUSY STATE & LOADER
        // ============================================================
        setBusy(false);
        TopLoader.forceClear();

        console.log("✅ Page initialization complete!");
    });

})(window, document);
