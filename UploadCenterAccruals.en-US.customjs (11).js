(function (w, d) {
  "use strict";

  /* =========================================================
   * ABSOLUTE INIT GUARD
   * ========================================================= */
  if (w.__DCL_ATTACH_INIT_DONE__) return;
  w.__DCL_ATTACH_INIT_DONE__ = true;

  /* =========================================================
   * CONFIG
   * ========================================================= */
  const FLOW_URL = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/41a79329e87f400fa632ea4e374e8eb0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EjqKzb2Ezk_4WSJ6yxrLA61AjOOlwvy7Y9usBb66K94";
  const DELETE_FLOW_URL = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/4a8a791d257b4953b9eb3461176bbca0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MaqUn5Iz6DwpE5ksq_cfydBBdOqkmSorNov_nMKmabs";
  const SYSINV_EXTRACT_FLOW_URL = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a30fd12c4ca34488b589d3261b5a2eba/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=X3vQ-GwyonKHmo-ZImPVA45_eYWelJgKYQU_tR6Vm2Y";

  const DCL_API_BASE = "/_api/cr650_dcl_masters";
  const DCL_SELECT = "$select=cr650_dclnumber,cr650_dcl_masterid,cr650_status"; const DOC_MASTER_API = "/_api/cr650_documents_masters?$select=cr650_documentname,cr650_documents_masterid";
  const DOCS_API_BASE = "/_api/cr650_dcl_documents";
  const DOCS_SELECT = "$select=cr650_doc_type,cr650_chargeamount,cr650_currencycode,cr650_documenturl,cr650_remarks,cr650_dcl_documentid,_cr650_dcl_number_value,cr650_file_extention,cr650_documentlanguage";

  // NEW: field on DCL master to store "docID;num1,num2|docID2;num3"
  const SYSTEM_INVOICE_FIELD = "cr650_sys_invoice";

  const DOC_CACHE_KEY = "dcl_doc_options_v2";
  const DOC_CACHE_TTL_MS = 12 * 60 * 60 * 1000;
  const CURR_CACHE_KEY = "dcl_currencies_v2";
  const CURR_CACHE_TTL = 24 * 60 * 60 * 1000;
  let DCL_STATUS = null;

  /* =========================================================
   * DOCUMENT TYPE HANDLING WITH AUTO-NUMBERING
   * ========================================================= */
  let UNIQUE_TYPES = new Set(); // now only used for "Other Documents"
  const norm = s => String(s || "").trim().toLowerCase();
  // NEW: strip trailing " 2", " 3", etc from a doc type
  function stripTypeSuffix(label) {
    return String(label || "").replace(/\s+\d+$/, "").trim();
  }
  // Track document type counts for automatic suffixing
  const docTypeCounters = new Map();

  function rebuildUniqueTypes() {
    // Only "Other Documents" is unique
    UNIQUE_TYPES = new Set(["other documents", "other document", "other"]);
  }

  function getDocumentTypeSuffix(docType) {
    const normalized = norm(docType);
    const count = docTypeCounters.get(normalized) || 0;
    docTypeCounters.set(normalized, count + 1);
    return count === 0 ? "" : ` ${count + 1}`;
  }

  function resetDocTypeCounters() {
    docTypeCounters.clear();
  }

  function recalculateDocTypeCounters() {
    resetDocTypeCounters();

    getAllRows().forEach(row => {
      // Only count existing saved rows, not new unsaved rows
      if (row.dataset.mode !== "existing") return;

      const typeText = row.querySelector('input[name="docType"]');
      if (!typeText || !typeText.value) return;

      const fullName = typeText.value.trim();

      // Parse the document name to extract base name and number
      // Matches patterns like "MOFA 2", "MOFA 3", etc.
      const match = fullName.match(/^(.+?)\s+(\d+)$/);

      const baseName = match ? match[1].trim() : fullName;
      const number = match ? parseInt(match[2], 10) : 1;

      // Track the highest number used for each base name
      const normalized = norm(baseName);
      const currentMax = docTypeCounters.get(normalized) || 0;
      docTypeCounters.set(normalized, Math.max(currentMax, number));
    });
  }

  /* =========================================================
   * TOP LOADER
   * ========================================================= */
  const Loading = (() => {
    let count = 0, timer = null;
    const bar = () => d.getElementById("top-loader");
    const scope = () => d.getElementById("dclWizard");
    function show() {
      const el = bar(); if (!el) return;
      el.classList.remove("hidden");
      el.setAttribute("aria-hidden", "false");
      const sc = scope(); if (sc) sc.classList.add("blur-while-loading");
    }
    function hide() {
      const el = bar(); if (!el) return;
      el.classList.add("hidden");
      el.setAttribute("aria-hidden", "true");
      const sc = scope(); if (sc) sc.classList.remove("blur-while-loading");
    }
    return {
      start() { count++; clearTimeout(timer); timer = setTimeout(() => { if (count > 0) show(); }, 80); },
      stop() { count = Math.max(0, count - 1); if (count === 0) { clearTimeout(timer); timer = setTimeout(hide, 120); } }
    };
  })();
  w.Loading = Loading;

  /* =========================================================
   * PORTALS TOKEN + safeAjax
   * ========================================================= */
  function getPortalTokenSync() {
    const input = d.querySelector('input[name="__RequestVerificationToken"]');
    if (input?.value) return input.value;
    const meta = d.querySelector('meta[name="__RequestVerificationToken"]');
    if (meta?.content) return meta.content;
    const c = (d.cookie || "").split("; ").find(x => x.startsWith("__RequestVerificationToken="));
    if (c) return decodeURIComponent(c.split("=")[1]);
    return "";
  }

  (function ensureSafeAjax(webapi, $) {
    if (webapi.safeAjax) return;

    if ($ && w.shell && typeof w.shell.getTokenDeferred === "function") {
      webapi.safeAjax = function (ajaxOptions) {
        const dfd = $.Deferred();
        w.shell.getTokenDeferred().done(function (token) {
          try {
            ajaxOptions = ajaxOptions || {};
            ajaxOptions.headers = ajaxOptions.headers || {};
            ajaxOptions.headers["__RequestVerificationToken"] = token;
            Loading.start();
            $.ajax(ajaxOptions)
              .done(function () { dfd.resolve.apply(dfd, arguments); })
              .fail(function () { dfd.reject.apply(dfd, arguments); })
              .always(function () { Loading.stop(); });
          } catch (e) {
            Loading.stop();
            dfd.reject(e);
          }
        }).fail(function () {
          dfd.rejectWith(this, arguments);
        });
        return dfd.promise();
      };
    } else {
      webapi.safeAjax = function (ajaxOptions) {
        const token = getPortalTokenSync();
        const method = (ajaxOptions.type || ajaxOptions.method || "GET").toUpperCase();
        const body = (method === "GET" || method === "HEAD") ? undefined : ajaxOptions.data;
        const headers = Object.assign({
          "Accept": "application/json",
          "Content-Type": ajaxOptions.contentType || "application/json; charset=utf-8",
          "__RequestVerificationToken": token,
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          "X-Requested-With": "XMLHttpRequest"
        }, ajaxOptions.headers || {});
        Loading.start();
        return fetch(ajaxOptions.url, { method, headers, body, credentials: "same-origin", cache: "no-store" })
          .then(async res => {
            if (!res.ok) {
              const t = await res.text().catch(() => "");
              const e = new Error("Request failed " + res.status + ": " + t);
              e.status = res.status; e.responseText = t; throw e;
            }
            if (res.status === 204) return {};
            const ct = res.headers.get("content-type") || "";
            return ct.includes("application/json") ? res.json() : res.text();
          })
          .finally(() => Loading.stop());
      };
    }
  })(w.webapi = w.webapi || {}, w.jQuery);

  /* =========================================================
   * UTILS
   * ========================================================= */
  const isGuid = s => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s || ""));
  const asGuid = g => String(g);
  const escapeHtml = s => String(s).replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
  const formatMoney = (n, currency) => {
    const val = (Number.isFinite(n) ? n : 0).toFixed(2);
    const symbols = { USD: "$", SAR: "﷼", AED: "د.إ", EUR: "€", GBP: "£", JPY: "¥" };
    const sym = currency ? (symbols[currency] || currency) : "$";
    return sym + " " + val;
  };

  function getIdFromUrl() {
    const u = new URL(w.location.href);
    for (const k of ["id", "dclid", "dcl_masterid", "record", "row", "masterid"]) {
      const v = (u.searchParams.get(k) || "").trim(); if (isGuid(v)) return v;
    }
    const m = u.href.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
    return m ? m[0] : "";
  }

  function errBox(where, err) {
    const status = err?.status || "n/a";
    const details = (err?.responseText || err?.message || "").slice(0, 800);
    console.error(`[${where}]`, err);
    alert(`[${where}] failed (status ${status}).\n${details}`);
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const res = reader.result || "";
        const base64 = res.split(",")[1] || "";
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async function sendToFlow(payload) {
    Loading.start();
    try {
      const res = await fetch(FLOW_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const text = await res.text();
      if (!res.ok) {
        const e = new Error("Flow call failed " + res.status + ": " + text);
        e.status = res.status; e.responseText = text; throw e;
      }
      try { return JSON.parse(text); } catch { return { success: true, raw: text }; }
    } finally {
      Loading.stop();
    }
  }

  // Call System Invoice extraction Flow with base64 file and extension
  async function extractSystemInvoiceNumbersFromFlow(fileBase64, fileExtension) {
    if (!fileBase64) return [];

    const ext = (fileExtension || "").replace(/^\./, "").toLowerCase() || "pdf";

    const body = {
      file_base64: fileBase64,
      file_extension: ext
    };

    // Show top-loader while talking to the extractor Flow
    Loading.start();
    try {
      const res = await fetch(SYSINV_EXTRACT_FLOW_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });

      const text = await res.text();

      if (!res.ok) {
        const err = new Error("System invoice extraction flow failed " + res.status + ": " + text);
        err.status = res.status;
        err.responseText = text;
        throw err;
      }

      let json;
      try {
        json = JSON.parse(text);
      } catch {
        return [];
      }

      const numbers = Array.isArray(json.numbers) ? json.numbers : [];
      return numbers
        .map(v => String(v).trim())
        .filter(Boolean);

    } finally {
      Loading.stop();
    }
  }



  function isCiOrPl(label) {
    const n = (label || "").trim().toLowerCase();
    return n === "packing list" || n === "commercial invoice" || n === "commerical invoice";
  }

  function isSystemInvoiceLabel(label) {
    // normalize and remove trailing number (e.g. "System Invoice 2" → "system invoice")
    const base = norm(label).replace(/\s+\d+$/, "");
    return base === "system invoice" || base === "system invoices";
  }



  // Prefer readonly URL for existing rows
  function getRowUrl(row) {
    const urlEl = row.querySelector('.field--fileurl .doc-url') || row.querySelector('.doc-url');
    const v = (urlEl && urlEl.value ? urlEl.value : "").trim();
    if (!v || v === "null" || v === "undefined") return "";
    return v;
  }

  const ALLOWED_EXTENSIONS = new Set(["docx", "xlsx", "png", "jpg", "jpeg", "pdf"]);

  function getFileExtension(name) {
    const parts = String(name || "").toLowerCase().split(".");
    return parts.length > 1 ? parts.pop() : "";
  }

  /* =========================================================
   * DATA LOADERS
   * ========================================================= */
  async function loadDclNumberInto(spanId = "dclNumber") {
    const el = d.getElementById(spanId);
    if (!el) return;

    const id = getIdFromUrl();
    if (!id) {
      el.textContent = "-";
      return;
    }

    try {
      const r = await webapi.safeAjax({
        type: "GET",
        url: `${DCL_API_BASE}(${id})?${DCL_SELECT}`,
        contentType: "application/json"
      });

      const num = r?.cr650_dclnumber || "-";
      el.textContent = num;
      w.__CURRENT_DCL_NUMBER = num;

      // ✅ NEW: Capture status
      DCL_STATUS = r?.cr650_status || null;

    } catch {
      try {
        const filter = `cr650_dcl_masterid eq ${asGuid(id)}`;
        const url = `${DCL_API_BASE}?${DCL_SELECT}&$filter=${filter}`;
        const r2 = await webapi.safeAjax({
          type: "GET",
          url,
          contentType: "application/json"
        });

        const num = r2?.value?.[0]?.cr650_dclnumber || "-";
        el.textContent = num;
        w.__CURRENT_DCL_NUMBER = num;

        // ✅ NEW: Capture status
        DCL_STATUS = r2?.value?.[0]?.cr650_status || null;

      } catch {
        el.textContent = "-";
        w.__CURRENT_DCL_NUMBER = "-";
        DCL_STATUS = null;
      }
    }
  }

  let DOC_OPTIONS = [];
  async function ensureDocOptionsLoaded() {
    try {
      const c = JSON.parse(localStorage.getItem(DOC_CACHE_KEY) || "null");
      if (c && Array.isArray(c.items) && (Date.now() - c.savedAt) < DOC_CACHE_TTL_MS) {
        DOC_OPTIONS = c.items;
        rebuildUniqueTypes();
        return;
      }
    } catch { }
    try {
      const r = await webapi.safeAjax({ type: "GET", url: DOC_MASTER_API, contentType: "application/json" });
      const items = Array.isArray(r?.value) ? r.value : [];
      DOC_OPTIONS = items.filter(i => i.cr650_documentname).map(i => ({
        value: i.cr650_documents_masterid, label: i.cr650_documentname
      }));
      if (!DOC_OPTIONS.some(o => /Other Documents/i.test(o.label)))
        DOC_OPTIONS.push({ value: "other-doc", label: "Other Documents" });
      rebuildUniqueTypes();
      localStorage.setItem(DOC_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), items: DOC_OPTIONS }));
    } catch {
      DOC_OPTIONS = [{ value: "other-doc", label: "Other Documents" }];
      rebuildUniqueTypes();
    }
  }

  /* =========================================================
   * COMPANY TYPE AND CURRENCY HANDLING
   * ========================================================= */
  let COMPANY_INFO = {
    type: null,
    defaultCurrency: "AED",
    currencyLabel: "AED",
    dclCurrency: null
  };

  async function loadCompanyInfo() {
    try {
      const id = getIdFromUrl();
      if (!id) return;

      const selectFields = "$select=cr650_company_type,cr650_currencycode,cr650_dcl_masterid";
      const url = `${DCL_API_BASE}(${id})?${selectFields}`;
      const r = await webapi.safeAjax({
        type: "GET",
        url: url,
        contentType: "application/json"
      });

      const companyType = String(r?.cr650_company_type ?? "").trim();
      const dclCurrency = String(r?.cr650_currencycode ?? "").trim();

      if (companyType === "0") {
        COMPANY_INFO = {
          type: "Technolube",
          defaultCurrency: "USD",
          currencyLabel: "USD",
          dclCurrency: dclCurrency || "USD"
        };
      } else if (companyType === "1") {
        COMPANY_INFO = {
          type: "Petrolube",
          defaultCurrency: "SAR",
          currencyLabel: "SAR",
          dclCurrency: dclCurrency || "SAR"
        };
      } else {
        COMPANY_INFO = {
          type: companyType ? `Unknown (${companyType})` : "Unknown",
          defaultCurrency: "AED",
          currencyLabel: "AED",
          dclCurrency: dclCurrency || "AED"
        };
      }

      console.log("Company info loaded:", COMPANY_INFO);

    } catch (e) {
      console.warn("Could not load company info, using default currency", e);
      COMPANY_INFO = { type: null, defaultCurrency: "AED", currencyLabel: "AED", dclCurrency: "AED" };
    }
  }

  /* =========================================================
   * CURRENCY CONVERSION (Fawaz Ahmed API)
   * ========================================================= */
  const conversionCache = new Map();

  async function convertCurrency(amount, fromCurrency, toCurrency) {
    if (!amount || amount === 0) return 0;
    if (fromCurrency === toCurrency) return amount;

    const cacheKey = `${fromCurrency}_${toCurrency}`;
    const cached = conversionCache.get(cacheKey);

    if (cached && (Date.now() - cached.timestamp) < 3600000) {
      return amount * cached.rate;
    }

    const fromLower = fromCurrency.toLowerCase();
    const toLower = toCurrency.toLowerCase();

    const primaryUrl = `https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/${fromLower}.json`;
    const fallbackUrl = `https://latest.currency-api.pages.dev/v1/currencies/${fromLower}.json`;

    try {
      let data = null;

      try {
        const response = await fetch(primaryUrl);
        if (!response.ok) throw new Error("Primary API failed");
        data = await response.json();
      } catch (primaryError) {
        console.warn("Primary currency API failed, trying fallback:", primaryError);
        const fallbackResponse = await fetch(fallbackUrl);
        if (!fallbackResponse.ok) throw new Error("Fallback API also failed");
        data = await fallbackResponse.json();
      }

      const rate = data?.[fromLower]?.[toLower];

      if (!rate || typeof rate !== "number") {
        throw new Error(`Rate not found for ${fromCurrency}→${toCurrency}`);
      }

      conversionCache.set(cacheKey, { rate, timestamp: Date.now() });

      return amount * rate;

    } catch (e) {
      console.error("Currency conversion failed:", e);

      const fallbackRates = {
        "usd_sar": 3.75,
        "sar_usd": 0.2667,
        "usd_aed": 3.67,
        "aed_usd": 0.2725,
        "sar_aed": 0.98,
        "aed_sar": 1.02,
        "usd_eur": 0.92,
        "eur_usd": 1.09,
        "usd_gbp": 0.79,
        "gbp_usd": 1.27
      };

      const fallbackKey = `${fromLower}_${toLower}`;
      if (fallbackRates[fallbackKey]) {
        console.warn(`Using fallback rate for ${fromCurrency}→${toCurrency}`);
        return amount * fallbackRates[fallbackKey];
      }

      return amount;
    }
  }

  let CURRENCIES = [];
  const currencySelects = new Set();
  async function ensureCurrenciesLoaded() {
    try {
      const c = JSON.parse(localStorage.getItem(CURR_CACHE_KEY) || "null");
      if (c && Array.isArray(c.items) && (Date.now() - c.savedAt) < CURR_CACHE_TTL) {
        CURRENCIES = c.items;
        return;
      }
    } catch { }

    try {
      const primaryUrl = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies.json";
      const fallbackUrl = "https://latest.currency-api.pages.dev/v1/currencies.json";

      let data = null;

      try {
        const response = await fetch(primaryUrl);
        if (!response.ok) throw new Error("Primary currencies API failed");
        data = await response.json();
      } catch (primaryError) {
        console.warn("Primary currencies API failed, trying fallback:", primaryError);
        const fallbackResponse = await fetch(fallbackUrl);
        if (!fallbackResponse.ok) throw new Error("Fallback currencies API also failed");
        data = await fallbackResponse.json();
      }

      CURRENCIES = Object.entries(data || {})
        .map(([code, name]) => ({
          code: code.toUpperCase(),
          name: String(name)
        }))
        .sort((a, b) => a.code.localeCompare(b.code));

      const priorityCodes = ["USD", "SAR", "EUR", "AED", "GBP"];
      const priority = [];
      const others = [];

      CURRENCIES.forEach(curr => {
        if (priorityCodes.includes(curr.code)) {
          priority.push(curr);
        } else {
          others.push(curr);
        }
      });

      CURRENCIES = [
        ...priorityCodes.map(code => priority.find(c => c.code === code)).filter(Boolean),
        ...others
      ];

      localStorage.setItem(CURR_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), items: CURRENCIES }));

    } catch (e) {
      console.error("Failed to load currencies from API, using fallback:", e);
      CURRENCIES = [
        { code: "USD", name: "United States Dollar" },
        { code: "SAR", name: "Saudi Riyal" },
        { code: "EUR", name: "Euro" },
        { code: "AED", name: "UAE Dirham" },
        { code: "GBP", name: "Pound Sterling" },
        { code: "JPY", name: "Japanese Yen" },
        { code: "CHF", name: "Swiss Franc" },
        { code: "CNY", name: "Chinese Yuan" }
      ];
    }
  }

  function populateCurrencySelect(sel) {
    if (!sel) return;
    const keep = sel.value || COMPANY_INFO.defaultCurrency;

    const priorityCodes = ["USD", "SAR", "EUR", "AED", "GBP"];
    const hasPriority = CURRENCIES.some(c => priorityCodes.includes(c.code));

    let options = [`<option value="">Currency</option>`];

    CURRENCIES.forEach((c, idx) => {
      if (hasPriority && idx > 0 && priorityCodes.includes(CURRENCIES[idx - 1].code) && !priorityCodes.includes(c.code)) {
        options.push(`<option disabled>──────────</option>`);
      }
      options.push(`<option value="${c.code}">${c.code} — ${escapeHtml(c.name)}</option>`);
    });

    sel.innerHTML = options.join("");
    if (keep) sel.value = keep;
  }

  async function loadExistingDocumentsFor(masterId) {
    try {
      const filter = `_cr650_dcl_number_value eq ${asGuid(masterId)}`;
      const url = `${DOCS_API_BASE}?${DOCS_SELECT}&$filter=${filter}`;
      const res = await webapi.safeAjax({ type: "GET", url, contentType: "application/json" });
      return Array.isArray(res?.value) ? res.value : [];
    } catch (e) {
      errBox("Load documents", e);
      return [];
    }
  }

  async function createDocument(masterId, payload) {
    const data = Object.assign({}, payload, {
      "cr650_dcl_number@odata.bind": `/cr650_dcl_masters(${masterId})`
    });
    return await new Promise((resolve, reject) => {
      webapi.safeAjax({
        type: "POST",
        url: DOCS_API_BASE,
        contentType: "application/json; charset=utf-8",
        headers: { "OData-MaxVersion": "4.0", "OData-Version": "4.0" },
        data: JSON.stringify(data),
        success: function (res, status, xhr) {
          const body = res || {};
          const idFromBody = body.cr650_dcl_documentid || body.id;
          const idFromHdr = (xhr && typeof xhr.getResponseHeader === "function") ? xhr.getResponseHeader("entityid") : null;
          resolve(Object.assign({}, body, { cr650_dcl_documentid: idFromBody || idFromHdr }));
        },
        error: reject
      });
    });
  }

  async function updateDocument(docId, payload) {
    return await new Promise((resolve, reject) => {
      webapi.safeAjax({
        type: "PATCH",
        url: `${DOCS_API_BASE}(${docId})`,
        contentType: "application/json; charset=utf-8",
        headers: { "If-Match": "*", "OData-MaxVersion": "4.0", "OData-Version": "4.0" },
        data: JSON.stringify(payload),
        success: resolve,
        error: reject
      });
    });
  }

  async function deleteDocument(docId) {
    return await new Promise((resolve, reject) => {
      webapi.safeAjax({
        type: "DELETE",
        url: `${DOCS_API_BASE}(${docId})`,
        headers: { "If-Match": "*", "OData-MaxVersion": "4.0", "OData-Version": "4.0" },
        success: resolve,
        error: reject
      });
    });
  }

  /* =========================================================
   * SYSTEM INVOICE MAP on DCL MASTER
   * ========================================================= */
  let SYSTEM_INVOICE_MAP = new Map();        // docId -> [invoice1, invoice2, ...]
  let SYSTEM_INVOICE_LOADED_FOR = null;     // masterId

  function parseSystemInvoiceString(raw) {
    const map = new Map();
    if (!raw) return map;

    String(raw)
      .split("|")
      .map(seg => seg.trim())
      .filter(Boolean)
      .forEach(segment => {
        // split ONLY on the first semicolon
        const match = segment.match(/^([^;]+);(.*)$/);
        if (!match) return;

        const docId = match[1].trim();
        const numsPart = match[2].trim();
        if (!docId || !numsPart) return;

        const invoices = numsPart
          .split(",")
          .map(v => v.trim())
          .filter(Boolean);

        if (invoices.length) {
          map.set(docId, invoices);
        }
      });

    return map;
  }

  function serializeSystemInvoiceMap(map) {
    if (!(map instanceof Map)) return "";

    const chunks = [];

    for (const [docId, list] of map.entries()) {
      const cleanId = String(docId).trim();
      if (!cleanId) continue;

      const cleanedValues = (Array.isArray(list) ? list : [])
        .map(v => String(v)
          .replace(/[|;]/g, " ")   // avoid breaking our “docId;nums|docId2;nums2” format
          .trim()
        )
        .filter(Boolean);

      if (!cleanedValues.length) continue;

      chunks.push(`${cleanId};${cleanedValues.join(",")}`);
    }

    return chunks.join("|");
  }


  async function ensureSystemInvoiceMapLoaded(masterId) {
    if (!masterId) return;
    if (SYSTEM_INVOICE_LOADED_FOR === masterId) return;

    try {
      const url = `${DCL_API_BASE}(${masterId})?$select=${SYSTEM_INVOICE_FIELD}`;
      const r = await webapi.safeAjax({
        type: "GET",
        url,
        contentType: "application/json; charset=utf-8"
      });
      const raw = r && r[SYSTEM_INVOICE_FIELD] ? r[SYSTEM_INVOICE_FIELD] : "";
      SYSTEM_INVOICE_MAP = parseSystemInvoiceString(raw);
      SYSTEM_INVOICE_LOADED_FOR = masterId;
    } catch (e) {
      console.warn("Could not load system invoice map", e);
      SYSTEM_INVOICE_MAP = new Map();
      SYSTEM_INVOICE_LOADED_FOR = masterId;
    }
  }
  async function upsertSystemInvoiceEntry(masterId, docId, invoices) {
    if (!masterId || !docId) return;

    // make sure we start from the latest map from Dataverse
    await ensureSystemInvoiceMapLoaded(masterId);

    const cleaned = Array.isArray(invoices)
      ? invoices
        .map(v => String(v).trim())
        .filter(Boolean)
      : [];

    if (!cleaned.length) {
      // no invoices = remove this doc from the map
      SYSTEM_INVOICE_MAP.delete(docId);
    } else {
      SYSTEM_INVOICE_MAP.set(docId, cleaned);
    }

    const serialized = serializeSystemInvoiceMap(SYSTEM_INVOICE_MAP);

    await new Promise((resolve, reject) => {
      webapi.safeAjax({
        type: "PATCH",
        url: `${DCL_API_BASE}(${masterId})`,
        contentType: "application/json; charset=utf-8",
        headers: {
          "If-Match": "*",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0"
        },
        data: JSON.stringify({ [SYSTEM_INVOICE_FIELD]: serialized }),
        success: resolve,
        error: reject
      });
    });
  }


  /* =========================================================
   * UNIQUE TYPE ENFORCEMENT HELPERS
   * ========================================================= */
  function getAllRows() {
    return Array.from(d.querySelectorAll(".attachment-row"));
  }

  function anyExistingRowHasType(labelNorm) {
    return getAllRows().some(row => {
      if (row.dataset.mode !== "existing") return false;
      const sel = row.querySelector('select[name="docType"]');
      const typeText = row.querySelector('input[name="docType"]');
      const txt = sel ? sel.options[sel.selectedIndex]?.textContent || "" : (typeText ? typeText.value || "" : "");
      return norm(txt) === labelNorm;
    });
  }

  function collectSelectedUniqueTypes(ignoreRow = null) {
    const set = new Set();
    getAllRows().forEach(row => {
      if (ignoreRow && row === ignoreRow) return;
      const sel = row.querySelector('select[name="docType"]');
      const typeText = row.querySelector('input[name="docType"]');
      const txt = sel ? sel.options[sel.selectedIndex]?.textContent || "" : (typeText ? typeText.value || "" : "");
      const n = norm(txt);
      if (UNIQUE_TYPES.has(n)) {
        if (row.dataset.mode === "existing" || (row.dataset.mode === "new" && n)) set.add(n);
      }
    });
    return set;
  }

  function refreshUniqueTypeLocks() {
    const taken = collectSelectedUniqueTypes();
    getAllRows().forEach(row => {
      const sel = row.querySelector('select[name="docType"]');
      if (!sel) return;
      const currentTxt = sel.options[sel.selectedIndex]?.textContent || "";
      const currentNorm = norm(currentTxt);

      Array.from(sel.options).forEach(opt => {
        const optNorm = norm(opt.textContent || "");
        if (!UNIQUE_TYPES.has(optNorm)) {
          opt.disabled = false;
          opt.removeAttribute("title");
          return;
        }
        const alreadyTaken = taken.has(optNorm);
        const thisRowHolds = (optNorm === currentNorm);
        opt.disabled = alreadyTaken && !thisRowHolds;
        if (opt.disabled) opt.title = "This document type already exists in this DCL";
        else opt.removeAttribute("title");
      });
    });
  }

  /* =========================================================
   * STATIC FIELD LOCKER
   * ========================================================= */
  function lockRowStaticFields(row) {
    const typeSel = row.querySelector('select[name="docType"]');
    if (typeSel) {
      typeSel.setAttribute("disabled", "disabled");
      typeSel.setAttribute("data-static", "1");
      typeSel.setAttribute("title", "Locked for existing records");
    }
    const urlInput = row.querySelector('.field--fileurl .doc-url');
    if (urlInput) {
      urlInput.readOnly = true;
      urlInput.setAttribute("aria-readonly", "true");
      urlInput.setAttribute("title", "Locked for existing records");
    }
  }

  /* =========================================================
   * UI BUILDERS
   * ========================================================= */
  let __rowIndexCounter = 0;
  function buildDocTypeOptions(selectedLabel) {
    const seen = new Set();
    const selectedBase = stripTypeSuffix(selectedLabel || "").toLowerCase();

    const opts = ['<option value="">Select Type</option>'];

    DOC_OPTIONS.forEach(o => {
      const base = stripTypeSuffix(o.label);
      if (!base) return;

      const key = base.toLowerCase();
      if (seen.has(key)) return;     // skip MOFA 2, MOFA 3 etc
      seen.add(key);

      const isSelected = key === selectedBase;
      opts.push(
        `<option value="${escapeHtml(base)}"${isSelected ? " selected" : ""}>${escapeHtml(base)}</option>`
      );
    });

    return opts.join("");
  }

  function buildAttachmentRow(existing) {
    const idx = __rowIndexCounter++;
    const isExisting = !!existing;
    const row = d.createElement("div");
    row.className = "attachment-row";
    row.dataset.mode = isExisting ? "existing" : "new";
    if (isExisting && existing.cr650_dcl_documentid) row.dataset.docId = existing.cr650_dcl_documentid;

    if (isExisting && existing.cr650_file_extention) {
      row.dataset.fileExt = String(existing.cr650_file_extention).toLowerCase();
    }
    if (isExisting && typeof existing.cr650_documentlanguage === "string" && existing.cr650_documentlanguage.trim()) {
      row.dataset.lang = existing.cr650_documentlanguage.trim();
    }

    const docLabel = existing?.cr650_doc_type ? String(existing.cr650_doc_type) : "";
    if (docLabel) {
      const baseLabel = stripTypeSuffix(docLabel);
      if (baseLabel && !DOC_OPTIONS.some(o => stripTypeSuffix(o.label) === baseLabel)) {
        DOC_OPTIONS.push({ value: "adhoc-" + baseLabel, label: baseLabel });
        rebuildUniqueTypes();
      }
    }


    row.innerHTML = `
      <button type="button" class="row-remove" title="${isExisting ? "Delete" : "Remove"}">${isExisting ? "🗑" : "×"}</button>
      <div class="row-grid">
        ${isExisting ? `
          <div class="field">
            <label for="docType_${idx}">Document Type</label>
            <input id="docType_${idx}" 
                   name="docType" 
                   type="text" 
                   class="ctrl doc-type-text" 
                   value="${docLabel ? escapeHtml(docLabel) : ""}" 
                   readonly 
                   aria-readonly="true" 
                   title="Document type (read-only)">
          </div>
        ` : `
          <div class="field">
            <label for="docType_${idx}">Document Type</label>
            <select id="docType_${idx}" name="docType" class="ctrl" required>
  ${buildDocTypeOptions(docLabel)}
</select>

          </div>
        `}

        <div class="field field--customfolder" style="display: none;">
          <label for="customFolder_${idx}">Custom Folder Name</label>
          <input id="customFolder_${idx}" type="text" class="ctrl custom-folder-name" 
                 placeholder="Enter folder name (e.g., Special Permits)" 
                 maxlength="100">
          <small class="field-hint">This name will be used for the SharePoint folder</small>
        </div>

        ${isExisting ? `
          <div class="field field--fileurl">
            <label for="docUrl_${idx}">Document URL</label>
            <div class="fileurl-input">
              <input id="docUrl_${idx}" type="url" class="ctrl doc-url" placeholder="https://..." value="${existing?.cr650_documenturl ? escapeHtml(existing.cr650_documenturl) : ""}" readonly aria-readonly="true" title="Locked for existing records">
              <a class="url-open" href="${existing?.cr650_documenturl ? escapeHtml(existing.cr650_documenturl) : "#"}" target="_blank" rel="noopener" ${existing?.cr650_documenturl ? "" : "aria-disabled='true'"}>Open</a>
            </div>
          </div>
        ` : `
          <div class="field field--fileupload">
            <label for="upload_${idx}">Upload document</label>
            <input id="upload_${idx}" type="file" class="ctrl doc-file" accept=".pdf,.docx,.xlsx,.png,.jpg,.jpeg">
            <input id="docUrl_${idx}" type="hidden" class="doc-url" value="">
          </div>
        `}

        <div class="field">
          <label for="charge_${idx}">Charge</label>
          <input id="charge_${idx}" type="number" class="ctrl" step="0.01" min="0" placeholder="Charge" value="${existing?.cr650_chargeamount ?? ""}">
        </div>

        <div class="field">
          <label for="currency_${idx}">Currency</label>
          <select id="currency_${idx}" class="ctrl currency-select">
            <option value="">Currency</option>
          </select>
        </div>

        <div class="field full field--remarks">
          <label for="remarks_${idx}">Remarks</label>
          <textarea id="remarks_${idx}" class="remarks" placeholder="Enter remarks for this file...">${existing?.cr650_remarks ? escapeHtml(existing.cr650_remarks) : ""}</textarea>
        </div>

        <!-- NEW: System Invoice numbers combo -->
        <div class="field full field--system-invoice" style="display:none;">
          <label for="sysInvInput_${idx}">System Invoice Numbers</label>
          <div class="sysinv-combo">
            <input id="sysInvInput_${idx}" type="text" class="ctrl sysinv-input"
                   placeholder="Type invoice number and press Enter" />
          </div>
          <div class="sysinv-tags"></div>
          <input type="hidden" class="sysinv-hidden" value="">
          <small class="field-hint">You can add multiple numbers; they will be stored with this document.</small>
        </div>

        <div class="field full field--actions">
          <button type="button" class="btn row-save">${isExisting ? "Save changes" : "Save new"}</button>
        </div>
      </div>
    `;

    // refs
    const removeBtn = row.querySelector(".row-remove");
    const saveBtn = row.querySelector(".row-save");
    const typeSel = row.querySelector(`select#docType_${idx}`);
    const typeText = row.querySelector(`input#docType_${idx}`);
    const chargeInput = row.querySelector(`#charge_${idx}`);
    const curSel = row.querySelector(`#currency_${idx}`);
    const remarksEl = row.querySelector(`#remarks_${idx}`);
    const urlInput = row.querySelector(`#docUrl_${idx}`);
    const fileInput = row.querySelector(`#upload_${idx}`);
    const urlOpen = row.querySelector(".url-open");

    // NEW: System Invoice combo elements
    const sysInvField = row.querySelector(".field--system-invoice");
    const sysInvInput = row.querySelector(".sysinv-input");
    const sysInvTags = row.querySelector(".sysinv-tags");
    const sysInvHidden = row.querySelector(".sysinv-hidden");
    let sysInvExtractionInProgress = false;

    async function maybeExtractSystemInvoiceFromCurrentFile() {
      if (!fileInput || !fileInput.files || !fileInput.files[0]) return;

      // Get current document type label
      let currentLabel = "";
      if (typeSel) {
        currentLabel = typeSel.options[typeSel.selectedIndex]?.textContent?.trim() || "";
      } else if (typeText) {
        currentLabel = typeText.value?.trim() || "";
      }

      if (!isSystemInvoiceLabel(currentLabel)) return;

      const file = fileInput.files[0];
      const ext = getFileExtension(file.name) || "pdf";

      if (sysInvExtractionInProgress) return;
      sysInvExtractionInProgress = true;

      try {
        const base64 = await fileToBase64(file);

        // Call extractor Flow
        const numbers = await extractSystemInvoiceNumbersFromFlow(base64, ext);

        // Clear existing tags and re add
        clearSystemInvoiceTags();
        if (numbers.length) {
          if (sysInvField) {
            sysInvField.style.display = "block";
          }
          numbers.forEach(n => addSystemInvoiceTag(n));
        }
      } catch (e) {
        console.warn("System invoice numbers extraction failed:", e);
      } finally {
        sysInvExtractionInProgress = false;
      }
    }

    function syncSystemInvoiceHidden() {
      if (!sysInvHidden || !sysInvTags) return;
      const values = Array.from(sysInvTags.querySelectorAll(".sysinv-tag"))
        .map(tag => tag.dataset.value || tag.textContent || "")
        .map(v => v.trim())
        .filter(Boolean);
      sysInvHidden.value = values.join(",");
    }

    function addSystemInvoiceTag(value) {
      if (!sysInvTags || !value) return;
      const v = String(value).trim();
      if (!v) return;

      const exists = Array.from(sysInvTags.querySelectorAll(".sysinv-tag"))
        .some(tag => (tag.dataset.value || tag.textContent || "").trim().toLowerCase() === v.toLowerCase());
      if (exists) return;

      const span = d.createElement("span");
      span.className = "sysinv-tag";
      span.dataset.value = v;
      span.textContent = v;

      const btn = d.createElement("button");
      btn.type = "button";
      btn.className = "sysinv-tag-remove";
      btn.textContent = "×";
      btn.addEventListener("click", () => {
        span.remove();
        syncSystemInvoiceHidden();
      });

      span.appendChild(btn);
      sysInvTags.appendChild(span);
      syncSystemInvoiceHidden();
    }

    function clearSystemInvoiceTags() {
      if (!sysInvTags || !sysInvHidden) return;
      sysInvTags.innerHTML = "";
      sysInvHidden.value = "";
    }

    function getSystemInvoiceValuesForRow() {
      if (!sysInvHidden) return [];
      return String(sysInvHidden.value || "")
        .split(",")
        .map(v => v.trim())
        .filter(Boolean);
    }

    function updateSystemInvoiceVisibilityByLabel(label) {
      if (!sysInvField) return;
      const isSys = isSystemInvoiceLabel(label || "");
      sysInvField.style.display = isSys ? "block" : "none";
      if (!isSys) {
        clearSystemInvoiceTags();
      }
    }

    if (sysInvInput) {
      sysInvInput.addEventListener("keydown", e => {
        if (e.key === "Enter" || e.key === ",") {
          e.preventDefault();
          const v = sysInvInput.value;
          addSystemInvoiceTag(v);
          sysInvInput.value = "";
        }
      });
    }

    // For existing rows, display with suffix if present and preload System Invoice tags
    if (isExisting && typeText) {
      const currentValue = typeText.value || "";
      if (isCiOrPl(currentValue) && row.dataset.lang) {
        typeText.value = currentValue + " \u2013 " + row.dataset.lang;
      }

      if (isSystemInvoiceLabel(currentValue)) {
        updateSystemInvoiceVisibilityByLabel(currentValue);

        const docId = row.dataset.docId;
        if (docId && SYSTEM_INVOICE_MAP instanceof Map) {
          const vals = SYSTEM_INVOICE_MAP.get(docId) || [];
          vals.forEach(v => addSystemInvoiceTag(v));
        }
      }
    }

    // currencies
    currencySelects.add(curSel);
    populateCurrencySelect(curSel);
    if (existing?.cr650_currencycode) curSel.value = existing.cr650_currencycode;

    if (fileInput) {
      fileInput.addEventListener("change", () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) {
          if (urlInput) urlInput.value = "";
          clearSystemInvoiceTags();
          return;
        }
        const ext = getFileExtension(file.name);
        if (!ALLOWED_EXTENSIONS.has(ext)) {
          alert("Invalid file type. Allowed types are: DOCX, XLSX, PNG, JPG, JPEG, PDF.");
          fileInput.value = "";
          if (urlInput) urlInput.value = "";
          clearSystemInvoiceTags();
          return;
        }
        if (urlInput) urlInput.value = "uploaded:" + file.name;

        // NEW: if doc type is System Invoice, immediately request extraction
        maybeExtractSystemInvoiceFromCurrentFile();
      });
    }


    if (urlInput && urlOpen) {
      const syncUrl = () => {
        const u = (urlInput.value || "").trim();
        if (u) { urlOpen.setAttribute("href", u); urlOpen.removeAttribute("aria-disabled"); }
        else { urlOpen.setAttribute("href", "#"); urlOpen.setAttribute("aria-disabled", "true"); }
      };
      urlInput.addEventListener("input", syncUrl);
      syncUrl();
    }

    // cost + unique locks on change with debouncing
    let updateTimer = null;

    if (chargeInput) {
      chargeInput.addEventListener("input", () => {
        clearTimeout(updateTimer);
        updateTimer = setTimeout(() => {
          notifyCostChanged(row);
        }, 300);
      });

      chargeInput.addEventListener("blur", () => {
        clearTimeout(updateTimer);
        notifyCostChanged(row);
      });

      chargeInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          clearTimeout(updateTimer);
          e.target.blur();
        }
      });
    }

    if (curSel) {
      curSel.addEventListener("change", () => {
        clearTimeout(updateTimer);
        notifyCostChanged(row);
        refreshUniqueTypeLocks();
      });
    }

    if (remarksEl) {
      remarksEl.addEventListener("input", () => {
        clearTimeout(updateTimer);
        updateTimer = setTimeout(() => {
          notifyCostChanged(row);
        }, 500);
      });
    }

    if (typeSel) {
      function handleTypeChange() {
        clearTimeout(updateTimer);
        const selectedLabel = typeSel.options[typeSel.selectedIndex]?.textContent?.trim() || "";

        notifyCostChanged(row);
        refreshUniqueTypeLocks();
        updateSystemInvoiceVisibilityByLabel(selectedLabel);

        if (!isExisting) {
          const customFolderField = row.querySelector(".field--customfolder");
          const customFolderInput = row.querySelector(".custom-folder-name");
          const isOtherDoc = norm(selectedLabel).includes("other");

          if (customFolderField) {
            customFolderField.style.display = isOtherDoc ? "block" : "none";
          }
          if (customFolderInput && !isOtherDoc) {
            customFolderInput.value = "";
          }
        }

        // NEW: if the user just changed the type to System Invoice and a file is already uploaded, extract
        if (isSystemInvoiceLabel(selectedLabel)) {
          maybeExtractSystemInvoiceFromCurrentFile();
        } else {
          // if type is no longer System Invoice, clear tags
          clearSystemInvoiceTags();
        }
      }

      typeSel.addEventListener("change", handleTypeChange);

      const initialLabel = typeSel.options[typeSel.selectedIndex]?.textContent?.trim() || "";
      updateSystemInvoiceVisibilityByLabel(initialLabel);
    }


    // SAVE (create/update)
    saveBtn.addEventListener("click", async (ev) => {
      ev.preventDefault();
      if (row.__saving__) return;
      row.__saving__ = true;
      saveBtn.disabled = true;

      let selectedLabel = "";
      let selectedNorm = "";

      if (typeSel) {
        selectedLabel = typeSel.options[typeSel.selectedIndex]?.textContent?.trim() || "";
        selectedNorm = norm(selectedLabel);
      } else if (typeText) {
        selectedLabel = typeText.value?.trim() || "";
        selectedNorm = norm(selectedLabel);
      }

      const systemInvoiceValues = getSystemInvoiceValuesForRow();

      const basePayload = {
        cr650_chargeamount: chargeInput.value ? Number(chargeInput.value) : null,
        cr650_currencycode: curSel.value || null,
        cr650_remarks: (remarksEl.value || "").trim() || null
      };

      try {
        Loading.start();

        if (row.dataset.mode === "existing" && row.dataset.docId) {
          // UPDATE — type/url immutable
          await updateDocument(row.dataset.docId, basePayload);

          // Update system invoice mapping if applicable
          if (isSystemInvoiceLabel(selectedLabel)) {
            const masterId = getIdFromUrl();
            if (masterId) {
              await upsertSystemInvoiceEntry(masterId, row.dataset.docId, systemInvoiceValues);
            }
          }

          row.classList.add("row-saved");
          setTimeout(() => row.classList.remove("row-saved"), 800);

        } else {
          // CREATE — only check uniqueness for "Other Documents"
          if (UNIQUE_TYPES.has(selectedNorm)) {
            const customFolderInput = row.querySelector(".custom-folder-name");
            const customFolderName = customFolderInput ? (customFolderInput.value || "").trim() : "";

            if (!customFolderName) {
              if (anyExistingRowHasType(selectedNorm)) {
                throw new Error(`You already have a document of type "${selectedLabel}". Please enter a custom folder name or choose a different type.`);
              }
              const taken = collectSelectedUniqueTypes(row);
              if (taken.has(selectedNorm)) {
                throw new Error(`"${selectedLabel}" is already selected in another row. Please enter a custom folder name.`);
              }
            }
          }

          const masterId = getIdFromUrl();
          if (!masterId) throw new Error("Missing DCL master id in URL.");

          const file = fileInput && fileInput.files && fileInput.files[0];
          if (!file) throw new Error("Please upload a document file before saving.");
          if (!selectedLabel) throw new Error("Please select Document Type.");

          const chargeValue = chargeInput.value ? Number(chargeInput.value) : 0;
          if (chargeValue > 0 && !curSel.value) {
            throw new Error("Please select a currency when entering a charge amount.");
          }

          // Get custom folder name for "Other Documents"
          const customFolderInput = row.querySelector(".custom-folder-name");
          let customFolderName = "";

          if (norm(selectedLabel).includes("other") && customFolderInput) {
            customFolderName = (customFolderInput.value || "").trim();
            if (!customFolderName) {
              throw new Error("Please enter a custom folder name for 'Other Documents'.");
            }
            if (!/^[a-zA-Z0-9\s\-_()]+$/.test(customFolderName)) {
              throw new Error("Folder name can only contain letters, numbers, spaces, hyphens, underscores, and parentheses.");
            }
          }

          const ext = getFileExtension(file.name) || "docx";
          if (!ALLOWED_EXTENSIONS.has(ext)) {
            throw new Error("Invalid file type. Allowed types are: DOCX, XLSX, PNG, JPG, JPEG, PDF.");
          }

          // Build document type with auto-numbering
          let finalTypeName = selectedLabel;
          if (norm(selectedLabel).includes("other")) {
            finalTypeName = customFolderName.trim();
          } else {
            const baseName = selectedLabel.replace(/\s+\d+$/, "").trim();
            const suffix = getDocumentTypeSuffix(baseName);
            finalTypeName = baseName + suffix;
          }

          const createPayload = Object.assign({}, basePayload, {
            cr650_doc_type: finalTypeName,
            cr650_file_extention: ext
          });

          // 1) Create Dataverse row
          const created = await createDocument(masterId, createPayload);
          const newId = created?.cr650_dcl_documentid || created?.id;
          if (!newId) throw new Error("Created but id not returned. Check table permissions.");

          // 2) Upload file via main upload Flow
          const base64 = await fileToBase64(file);
          const flowBody = {
            id: masterId,
            fileContent: base64,
            docType: finalTypeName,
            fileExtension: ext,
            language: "en"
          };

          await sendToFlow(flowBody);



          // 3) Refresh to get URL and language
          let refreshed = null, updatedUrl = "";

          try {
            refreshed = await webapi.safeAjax({
              type: "GET",
              url: `${DOCS_API_BASE}(${newId})?${DOCS_SELECT}`,
              contentType: "application/json; charset=utf-8"
            });
            updatedUrl = refreshed?.cr650_documenturl || "";
          } catch { }

          if (refreshed && typeof refreshed.cr650_documentlanguage === "string" && refreshed.cr650_documentlanguage.trim()) {
            row.dataset.lang = refreshed.cr650_documentlanguage.trim();
          }

          // 4) Save System Invoice numbers if relevant (after auto-fill from Flow)
          if (isSystemInvoiceLabel(finalTypeName)) {
            const systemInvoiceValues = getSystemInvoiceValuesForRow(); // reads from hidden input populated by tags
            await upsertSystemInvoiceEntry(masterId, newId, systemInvoiceValues);
          }


          // Switch upload field → readonly URL
          const fileField = row.querySelector(".field--fileupload");
          if (fileField) {
            fileField.classList.remove("field--fileupload");
            fileField.classList.add("field--fileurl");
            fileField.innerHTML = `
              <label for="docUrl_${idx}">Document URL</label>
              <div class="fileurl-input">
                <input id="docUrl_${idx}" type="url" class="ctrl doc-url" value="${updatedUrl ? escapeHtml(updatedUrl) : ""}" readonly aria-readonly="true" title="Locked for existing records">
                <a class="url-open" href="${updatedUrl ? escapeHtml(updatedUrl) : "#"}" target="_blank" rel="noopener" ${updatedUrl ? "" : "aria-disabled='true'"}>Open</a>
              </div>
            `;
          }

          // Replace select with text input
          if (typeSel) {
            const typeField = typeSel.closest(".field");
            if (typeField) {
              typeField.innerHTML = `
                <label for="docType_${idx}">Document Type</label>
                <input id="docType_${idx}" 
                       name="docType" 
                       type="text" 
                       class="ctrl doc-type-text" 
                       value="${escapeHtml(finalTypeName)}" 
                       readonly 
                       aria-readonly="true" 
                       title="Document type (read-only)">
              `;
            }
          }

          // morph to existing
          row.dataset.mode = "existing";
          row.dataset.docId = newId;
          row.dataset.fileExt = ((refreshed?.cr650_file_extention || ext || "") + "").toLowerCase();
          row.querySelector(".row-save").textContent = "Save changes";
          row.querySelector(".row-remove").textContent = "🗑";
          row.querySelector(".row-remove").setAttribute("title", "Delete");

          const newTypeText = row.querySelector('input[name="docType"]');
          if (isCiOrPl(finalTypeName) && row.dataset.lang && newTypeText) {
            newTypeText.value = finalTypeName + " \u2013 " + row.dataset.lang;
          }

          row.classList.add("row-saved");
          setTimeout(() => row.classList.remove("row-saved"), 800);
        }

        notifyCostChanged(row);
        refreshUniqueTypeLocks();
        recalculateDocTypeCounters();

      } catch (err) {
        errBox("Save", err);
      } finally {
        Loading.stop();
        saveBtn.disabled = false;
        row.__saving__ = false;
      }
    });

    // DELETE
    removeBtn.addEventListener("click", async (ev) => {
      ev.preventDefault();

      const host = row.closest(".attachment-container") || d.getElementById("attachmentContainer");
      const costHost = d.getElementById("costCategories");
      const grandEl = d.getElementById("grandTotal");

      // New row, just remove from UI
      if (row.dataset.mode !== "existing" || !row.dataset.docId) {
        row.remove();
        currencySelects.delete(row.querySelector(".currency-select"));
        await updateCostBreakdown(host, costHost, grandEl);
        refreshUniqueTypeLocks();
        recalculateDocTypeCounters();
        return;
      }

      if (!confirm("Delete this document? This cannot be undone.")) return;

      try {
        const currentUrl = getRowUrl(row);
        const docId = row.dataset.docId;

        if (currentUrl) {
          const dclNumber = (w.__CURRENT_DCL_NUMBER || "").trim();
          if (!dclNumber) throw new Error("Cannot resolve DCL number. Refresh page and try again.");

          const typeText = row.querySelector('input[name="docType"]');
          const docTypeDisplay = typeText ? typeText.value.trim() : "";
          if (!docTypeDisplay) throw new Error("Cannot resolve document type for deletion.");

          let rawType = docTypeDisplay;
          let langFromLabel = "";
          ([" \u2013 ", " - ", " \u2014 "]).some(sep => {
            const i = docTypeDisplay.lastIndexOf(sep);
            if (i > -1) {
              rawType = docTypeDisplay.slice(0, i).trim();
              langFromLabel = docTypeDisplay.slice(i + sep.length).trim();
              return true;
            }
            return false;
          });

          const isLang = /^[A-Za-z]{2,10}(-[A-Za-z]{2,4})?$/.test(langFromLabel);
          const langForDelete = (row.dataset.lang || (isLang ? langFromLabel : "") || "").trim() || undefined;

          let fileExt = (row.dataset.fileExt || "").toLowerCase().trim();
          if ((!fileExt || fileExt === "null" || fileExt === "undefined") && isCiOrPl(rawType)) {
            fileExt = "docx";
          }

          if (isCiOrPl(rawType)) {
            await deleteFileViaFlow(dclNumber, rawType, fileExt, langForDelete);
          } else {
            await deleteFileViaFlow(dclNumber, rawType, fileExt);
          }
        }

        await deleteDocument(docId);

        // Remove System Invoice mapping if this was a System Invoice document
        const masterId = getIdFromUrl();
        if (masterId && isSystemInvoiceLabel(row.querySelector('input[name="docType"]')?.value || "")) {
          await upsertSystemInvoiceEntry(masterId, docId, []);
        }

        row.remove();
        currencySelects.delete(row.querySelector(".currency-select"));
        await updateCostBreakdown(host, costHost, grandEl);
        refreshUniqueTypeLocks();
        recalculateDocTypeCounters();

      } catch (err) {
        errBox("Delete", err);
      }
    });

    return row;
  }

  /* =========================================================
   * DELETE FILE VIA FLOW (with conditional language)
   * ========================================================= */
  async function deleteFileViaFlow(dclNumber, docType, fileExtension, language) {
    if (!dclNumber || !docType) throw new Error("Missing dclNumber or docType for deletion.");
    const body = { dclNumber, docType };
    if (fileExtension) body.fileExtension = String(fileExtension).replace(/^\./, "").toLowerCase();
    if (language && isCiOrPl(docType)) body.language = String(language).trim();

    Loading.start();
    try {
      const res = await fetch(DELETE_FLOW_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
      const text = await res.text();
      if (!res.ok) {
        const e = new Error("Delete-file flow failed " + res.status + ": " + text);
        e.status = res.status; e.responseText = text; throw e;
      }
      let json; try { json = JSON.parse(text); } catch { json = {}; }
      const ok = json && String(json.status).toLowerCase() === "success";
      if (!ok) throw new Error(json?.message || "Unexpected response from delete-file flow.");
      return true;
    } finally {
      Loading.stop();
    }
  }

  /* =========================================================
   * COST BREAKDOWN
   * ========================================================= */
  async function notifyCostChanged(row) {
    const host = row.closest(".attachment-container") || d.getElementById("attachmentContainer");
    const costHost = d.getElementById("costCategories");
    const grandEl = d.getElementById("grandTotal");
    await updateCostBreakdown(host, costHost, grandEl);
  }

  async function updateCostBreakdown(containerEl, costHostEl, grandTotalEl) {
    if (!containerEl || !costHostEl) return;
    const existingLoader = costHostEl.querySelector(".loading-indicator");
    costHostEl.querySelectorAll(".cost-category.dynamic, .cost-category.total:not(.loading-indicator)").forEach(n => n.remove());

    if (!existingLoader) {
      const loadingDiv = d.createElement("div");
      loadingDiv.className = "cost-category total loading-indicator";
      loadingDiv.innerHTML = `<span>Calculating totals...</span><span class="spinner">⏳</span>`;
      loadingDiv.style.cssText = "opacity: 1; transition: opacity 0.2s;";
      costHostEl.appendChild(loadingDiv);
    }

    const totals = new Map();
    const typeCounts = new Map();
    const charges = [];

    containerEl.querySelectorAll(".attachment-row").forEach(row => {
      const typeSel = row.querySelector('select[name="docType"]');
      const typeText = row.querySelector('input[name="docType"]');
      const chargeEl = row.querySelector('input[type="number"]');
      const currencyEl = row.querySelector(".currency-select");

      let label = "";
      if (typeSel && typeSel.selectedIndex > 0) {
        label = typeSel.options[typeSel.selectedIndex]?.textContent || "Uncategorized";
      } else if (typeText && typeText.value) {
        label = typeText.value || "Uncategorized";
      } else {
        label = "Uncategorized";
      }

      const baseLabel = label.replace(/\s+\d+$/, "").trim();
      const currentCount = (typeCounts.get(baseLabel) || 0) + 1;
      typeCounts.set(baseLabel, currentCount);

      if (currentCount > 1) {
        label = `${baseLabel} ${currentCount}`;
      } else {
        label = baseLabel;
      }

      const charge = chargeEl ? (parseFloat(chargeEl.value) || 0) : 0;
      const currency = currencyEl ? currencyEl.value : "";

      if (charge > 0) {
        totals.set(label, { amount: charge, currency: currency || "USD" });
        if (currency) {
          charges.push({ amount: charge, currency });
        }
      }
    });

    const frag = d.createDocumentFragment();

    for (const [label, data] of totals.entries()) {
      const div = d.createElement("div");
      div.className = "cost-category dynamic";
      const formattedAmount = data.amount.toFixed(2);
      div.innerHTML = `<span>${escapeHtml(label)}:</span><span>${formattedAmount} ${escapeHtml(data.currency)}</span>`;
      frag.appendChild(div);
    }

    costHostEl.appendChild(frag);

    const dclCurrency = COMPANY_INFO.dclCurrency || COMPANY_INFO.defaultCurrency;

    let totalInDclCurrency = 0;
    let totalInUsd = 0;
    let totalInSar = 0;
    let totalInAed = 0;

    try {
      for (const { amount, currency } of charges) {
        const inDcl = await convertCurrency(amount, currency, dclCurrency);
        totalInDclCurrency += inDcl;

        const inUsd = await convertCurrency(amount, currency, "USD");
        totalInUsd += inUsd;

        const inSar = await convertCurrency(amount, currency, "SAR");
        totalInSar += inSar;

        const inAed = await convertCurrency(amount, currency, "AED");
        totalInAed += inAed;
      }

      costHostEl.querySelectorAll(".loading-indicator").forEach(n => n.remove());

      // Always show DCL currency total
      const dclTotalDiv = d.createElement("div");
      dclTotalDiv.className = "cost-category total";
      dclTotalDiv.innerHTML =
        `<span>Total Accruals (${escapeHtml(dclCurrency)}):</span><span>${formatMoney(totalInDclCurrency, dclCurrency)}</span>`;
      costHostEl.appendChild(dclTotalDiv);

      // Always show USD total (skip if DCL currency is already USD)
      if (dclCurrency !== "USD") {
        const usdTotalDiv = d.createElement("div");
        usdTotalDiv.className = "cost-category total";
        usdTotalDiv.innerHTML =
          `<span>Total Accruals (USD):</span><span>${formatMoney(totalInUsd, "USD")}</span>`;
        costHostEl.appendChild(usdTotalDiv);
      }

      // Always show SAR total (skip if DCL currency is already SAR)
      if (dclCurrency !== "SAR") {
        const sarTotalDiv = d.createElement("div");
        sarTotalDiv.className = "cost-category total";
        sarTotalDiv.innerHTML =
          `<span>Total Accruals (SAR):</span><span>${formatMoney(totalInSar, "SAR")}</span>`;
        costHostEl.appendChild(sarTotalDiv);
      }

      // Always show AED total (skip if DCL currency is already AED)
      if (dclCurrency !== "AED") {
        const aedTotalDiv = d.createElement("div");
        aedTotalDiv.className = "cost-category total";
        aedTotalDiv.innerHTML =
          `<span>Total Accruals (AED):</span><span>${formatMoney(totalInAed, "AED")}</span>`;
        costHostEl.appendChild(aedTotalDiv);
      }

    } catch (e) {
      console.error("Currency conversion failed:", e);
      costHostEl.querySelectorAll(".loading-indicator").forEach(n => n.remove());

      let fallbackTotal = 0;
      for (const [, data] of totals.entries()) {
        fallbackTotal += data.amount;
      }

      const fallbackDiv = d.createElement("div");
      fallbackDiv.className = "cost-category total";
      fallbackDiv.innerHTML = `<span>Total Accruals (Mixed):</span><span>${formatMoney(fallbackTotal)}</span>`;
      costHostEl.appendChild(fallbackDiv);
    }
  }

  /* =========================================================
   * WIZARD LINK ID PROPAGATION + NAV LOADER
   * ========================================================= */
  function withSameId(href, id) {
    try {
      const u = new URL(href, window.location.origin);
      if (id) u.searchParams.set("id", id);
      return u.pathname + (u.search ? "?" + u.searchParams.toString() : "") + (u.hash || "");
    } catch {
      return href + (href.includes("?") ? "&" : "?") + "id=" + encodeURIComponent(id);
    }
  }

  function rewriteWizardLinks(id) {
    if (!id) return;
    const scope = document.getElementById("dclWizard") || document;
    const anchors = scope.querySelectorAll("#stepIndicators a, .navigation-buttons a");
    anchors.forEach(a => {
      const raw = a.getAttribute("href");
      if (!raw) return;
      a.setAttribute("href", withSameId(raw, id));
    });
  }

  function wireNavLoader() {
    const scope = document.getElementById("dclWizard") || document;
    const anchors = scope.querySelectorAll("#stepIndicators a, .navigation-buttons a");
    anchors.forEach(a => {
      a.addEventListener("click", () => { try { window.Loading && window.Loading.start(); } catch { } }, { capture: true });
    });
  }
  /* =========================================================
   * FORM LOCKING FOR SUBMITTED DCLs
   * ========================================================= */
  function lockFormIfSubmitted() {
    try {
      const status = (DCL_STATUS || "").toLowerCase();

      if (status !== "submitted") {
        console.log("📝 Form is editable - status:", DCL_STATUS || "none");
        return;
      }

      console.log("🔒 Locking form - DCL status is 'Submitted'");

      // 1. Disable all attachment row controls
      const rows = getAllRows();
      rows.forEach(row => {
        // Disable all inputs in the row
        row.querySelectorAll("input, textarea, select, button").forEach(el => {
          el.disabled = true;
          el.style.cursor = "not-allowed";
          el.style.opacity = "0.6";
          el.style.pointerEvents = "none";
        });

        // Make row visually "locked"
        row.style.opacity = "0.7";
        row.style.pointerEvents = "none";
      });

      // 2. Disable and hide the "Add Attachment" button
      const addBtn = document.querySelector(".add-attachment-btn") ||
        document.getElementById("addAttachmentBtn");
      if (addBtn) {
        addBtn.disabled = true;
        addBtn.style.display = "none";
      }

      // 3. Disable all action buttons (Save, Remove)
      document.querySelectorAll(".row-save, .row-remove").forEach(btn => {
        btn.disabled = true;
        btn.style.display = "none";
      });

      // 4. Disable file inputs
      document.querySelectorAll('input[type="file"]').forEach(input => {
        input.disabled = true;
        input.style.cursor = "not-allowed";
      });

      // 5. Disable currency selects
      currencySelects.forEach(sel => {
        if (sel) {
          sel.disabled = true;
          sel.style.cursor = "not-allowed";
        }
      });

      // 6. Make textareas read-only
      document.querySelectorAll(".remarks").forEach(textarea => {
        textarea.readOnly = true;
        textarea.style.cursor = "not-allowed";
        textarea.style.backgroundColor = "#f5f5f5";
      });

      // 7. Disable system invoice inputs
      document.querySelectorAll(".sysinv-input, .sysinv-tag-remove").forEach(el => {
        el.disabled = true;
        el.style.pointerEvents = "none";
      });

      // 8. Show locked banner
      showLockedBanner();

      console.log("✅ Attachments form fully locked");

    } catch (lockError) {
      console.error("❌ Error while locking form:", lockError);
    }
  }

  function showLockedBanner() {
    try {
      const wizard = document.getElementById("dclWizard");
      if (!wizard) {
        console.warn("Wizard container not found");
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
          This DCL has been submitted and attachments cannot be modified.
        </div>
      </div>
    `;

      const firstSection = wizard.querySelector(".step-content") ||
        wizard.querySelector(".attachment-container");
      if (firstSection) {
        wizard.insertBefore(banner, firstSection);
      } else {
        wizard.insertBefore(banner, wizard.firstChild);
      }
    } catch (bannerError) {
      console.error("Error creating locked banner:", bannerError);
    }
  }
  /* =========================================================
   * INITIALIZE
   * ========================================================= */
  async function initOnce() {
    const page = document.getElementById("dclWizard");
    if (!page) return;
    if (page.dataset.initDone === "1") return;
    page.dataset.initDone = "1";

    const host = document.querySelector(".attachment-container") || document.getElementById("attachmentContainer");
    const addBtn = document.querySelector(".add-attachment-btn") || document.getElementById("addAttachmentBtn");
    const costHost = document.getElementById("costCategories");
    const grandEl = document.getElementById("grandTotal");
    if (!host || !addBtn || !costHost || !grandEl) return;

    if (addBtn.tagName === "BUTTON" && !addBtn.getAttribute("type")) addBtn.setAttribute("type", "button");
    if (!addBtn.__boundAddHandler__) {
      addBtn.addEventListener("click", async function onAddClick(ev) {
        ev.preventDefault();
        const row = buildAttachmentRow(null);
        host.appendChild(row);
        await updateCostBreakdown(host, costHost, grandEl);
        refreshUniqueTypeLocks();
      });
      addBtn.__boundAddHandler__ = true;
    }

    try {
      Loading.start();

      await loadDclNumberInto("dclNumber");
      await ensureDocOptionsLoaded();
      await ensureCurrenciesLoaded();
      await loadCompanyInfo();
      rebuildUniqueTypes();

      const masterId = getIdFromUrl();
      if (masterId) {
        await ensureSystemInvoiceMapLoaded(masterId);
      }

      if (!host.dataset.loadedOnce && masterId) {
        const docs = await loadExistingDocumentsFor(masterId);
        for (const rec of docs) {
          if (rec.cr650_doc_type && !DOC_OPTIONS.some(o => o.label === rec.cr650_doc_type)) {
            DOC_OPTIONS.push({ value: "adhoc-" + rec.cr650_doc_type, label: rec.cr650_doc_type });
            rebuildUniqueTypes();
          }
          host.appendChild(buildAttachmentRow(rec));
        }
        host.dataset.loadedOnce = "1";
      }

      if (!host.querySelector(".attachment-row")) {
        host.appendChild(buildAttachmentRow(null));
      }

      host.querySelectorAll(".currency-select").forEach(populateCurrencySelect);

      recalculateDocTypeCounters();
      await updateCostBreakdown(host, costHost, grandEl);
      refreshUniqueTypeLocks();

      const currentId = getIdFromUrl();
      if (currentId) rewriteWizardLinks(currentId);
      wireNavLoader();
      // ✅ NEW: Lock form if submitted (MUST BE LAST)
      lockFormIfSubmitted();
    } catch (e) {
      errBox("Initialize", e);
    } finally {
      Loading.stop();
    }
  }

  if (d.readyState === "loading") d.addEventListener("DOMContentLoaded", initOnce, { once: true });
  else initOnce();

})(window, document);
