(function (w, d, $) {
  "use strict";

  /* =========================
     0) DEBUG TOGGLE & PANEL
     ========================= */
  const DEBUG = false;
  function ensureDebugPanel() {
    let host = d.getElementById("ciDebugPanel");
    if (host) return host;
    const card = d.querySelector('article.doc-card[aria-labelledby="ciTitle"]') || d.body;
    host = d.createElement("div");
    host.id = "ciDebugPanel";
    host.style.cssText = "margin-top:8px;font-family:monospace;font-size:12px;white-space:pre-wrap;border:1px solid #ddd;padding:8px;border-radius:6px;max-height:260px;overflow:auto;background:#fafafa;";
    host.setAttribute("aria-live", "polite");
    const header = d.createElement("div");
    header.innerHTML = "<strong>CI Debug</strong> (only visible in DEBUG=true)";
    host.appendChild(header);
    const body = d.createElement("div");
    body.id = "ciDebugBody";
    host.appendChild(body);
    (card.querySelector(".doc-actions") || card).after(host);
    return host;
  }
  function debugLog(...args) {
    if (!DEBUG) return;
    const panel = ensureDebugPanel();
    const body = panel.querySelector("#ciDebugBody");
    const msg = args.map(a => {
      if (a instanceof Error) return `Error: ${a.message}`;
      if (typeof a === "object") {
        try { return JSON.stringify(a, null, 2); } catch { return String(a); }
      }
      return String(a);
    }).join(" ");
    body.textContent += msg + "\n";
  }

  /* =========================
     1) TOP LOADER (shared)
     ========================= */
  const Loading = (() => {
    let count = 0, timer = null, safetyTimer = null;
    const bar = () => d.getElementById("top-loader");
    const scope = () => d.getElementById("dclWizard");
    function show() { const b = bar(); if (!b) return; b.classList.remove("hidden"); b.setAttribute("aria-hidden", "false"); scope()?.classList.add("blur-while-loading"); }
    function hide() { const b = bar(); if (!b) return; b.classList.add("hidden"); b.setAttribute("aria-hidden", "true"); scope()?.classList.remove("blur-while-loading"); }
    function start() {
      count++;
      clearTimeout(timer);
      timer = setTimeout(() => { if (count > 0) show(); }, 100);
      // Safety: auto-hide after 30s to prevent permanent stuck loading
      clearTimeout(safetyTimer);
      safetyTimer = setTimeout(() => { if (count > 0) { console.warn("Loading safety timeout – forcing hide (count was", count, ")"); count = 0; hide(); } }, 30000);
    }
    function stop() {
      count = Math.max(0, count - 1);
      if (count === 0) { clearTimeout(timer); clearTimeout(safetyTimer); timer = setTimeout(hide, 120); }
    }
    return { start, stop, _forceShow: show, _forceHide: hide };
  })();

  /* =========================
     2) CONFIG & FIELD MAPS
     ========================= */
  const API_URL_MASTER = "/_api/cr650_dcl_masters";
  const API_URL_CONTAINERS = "/_api/cr650_dcl_containers";
  const API_URL_ITEMS = "/_api/cr650_dcl_container_itemses";
  const API_URL_LOADPLANS = "/_api/cr650_dcl_loading_plans";
  const API_URL_BRANDS = "/_api/cr650_dcl_brands";
  const API_URL_TERMS = "/_api/cr650_dcl_terms_conditionses";
  const DOC_API = "/_api/cr650_dcl_documents";
  const API_URL_NOTIFY = "/_api/cr650_dcl_notify_parties";
  const AR_REPORT_API = "/_api/cr650_dcl_ar_reports";
  const SHIPPED_API = "/_api/cr650_dcl_shipped_orderses";
  const SHIPPED_FIELDS = [
    "cr650_order_number",
    "cr650_item_no",
    "cr650_delivery_note",
    "cr650_shipment_date",
    "createdon",
    "cr650_container_no",
    "cr650_seal_no"
  ];
  const HS_CODES_API = "/_api/cr650_hscodes";
  const HS_CODES_FIELDS = ["cr650_itemnumber", "cr650_hscode", "cr650_itemdescription", "cr650_category"];


  const FLOW_URL = "https://5d4ad4612f8beb7ead61b88cce63d5.4e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/41a79329e87f400fa632ea4e374e8eb0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EjqKzb2Ezk_4WSJ6yxrLA61AjOOlwvy7Y9usBb66K94";

  const MF = {
    id: "cr650_dcl_masterid",
    dcl: "cr650_dclnumber",
    ciNumber: "cr650_ci_number",  // ✅ ADD THIS
    autoNumberCI: "cr650_autonumber_ic",  // ✅ ADD THIS
    billTo: "cr650_billto",
    shipTo: "cr650_shiptolocation",
    incoterms: "cr650_incoterms",
    transportMode: "cr650_transportationmode",
    loadingPort: "cr650_loadingport",
    destinationPort: "cr650_destinationport",
    countryOfOrigin: "cr650_countryoforigin",
    currency: "cr650_currencycode",
    piNumber: "cr650_pinumber",
    paymentTerms: "cr650_paymentterms",
    poNumber: "cr650_ponumber",
    customerPoNumber: "cr650_po_customer_number",
    lcNumber: "cr650_lc_number",
    lcIssueDate: "cr650_lcissuedate",
    loadingDate: "cr650_loadingdate",
    party: "cr650_party",
    consignee: "cr650_consignee",
    applicant: "cr650_applicant",
    companyType: "cr650_company_type",
    country: "cr650_country",
    status: "cr650_status",
    descriptionGoods: "cr650_description_goods_services"


  };

  let DCL_STATUS = null;


  const CF = { id: "cr650_dcl_containerid", dclLkp: "_cr650_dcl_number_value", code: "cr650_id", type: "cr650_container_type", weight: "cr650_container_weight", totalPkgs: "cr650_total_packages_loaded", totalGross: "cr650_total_gross_weight_kg", totalNet: "cr650_total_net_weight_kg" };
  const IF = { id: "cr650_dcl_container_itemsid", containerLkp: "_cr650_dcl_number_value", planLkp: "_cr650_loadingplanitem_value", qty: "cr650_quantity" };
  const LF_CI = { id: "cr650_dcl_loading_planid", orderNo: "cr650_ordernumber", itemCode: "cr650_itemcode", hsCode: "cr650_hscode", description: "cr650_itemdescription", packaging: "cr650_packagingdetails", unitPrice: "cr650_unitprice", vat5: "cr650_vat5percent", totalInclVat: "cr650_totalincludingvat" };
  const LF_PL = {
    id: "cr650_dcl_loading_planid",
    orderNo: "cr650_ordernumber",
    itemCode: "cr650_itemcode",
    hsCode: "cr650_hscode",
    description: "cr650_itemdescription",
    packaging: "cr650_packagingdetails",
    uom: "cr650_unitofmeasure",
    totalLiters: "cr650_totalvolumeorweight",
    netKg: "cr650_netweightkg",
    grossKg: "cr650_grossweightkg",
    loadedQty: "cr650_loadedquantity",
    containerNumber: "cr650_containernumber",
    palletWeight: "cr650_palletweight"
  };
  const AR_LOGO_URL = "https://exportoperationsystem.powerappsportals.com/loading_plan_view/petrolube_arabic.jpg";
  const EN_LOGO_URL = "https://exportoperationsystem.powerappsportals.com/loading_plan_view/petrolube_english.jpg";
  const TECHNOLUBE_LOGO_URL = "https://exportoperationsystem.powerappsportals.com/loading_plan_view/technolube_logo.jpeg";
  const TECHNOLUBE_FOOTER_LOGO_URL = "https://exportoperationsystem.powerappsportals.com/loading_plan_view/Footer_technolube.jpg";

  // Preload all logo images as ArrayBuffers so they are ready for docx generation
  const ALL_LOGO_URLS = [AR_LOGO_URL, EN_LOGO_URL, TECHNOLUBE_LOGO_URL, TECHNOLUBE_FOOTER_LOGO_URL];
  const LOGO_CACHE = new Map();

  /** Fetch a single image with retries and multiple credential modes */
  async function fetchImageBuffer(url) {
    // Try different credential modes — portal may need "include" for same-origin
    for (const creds of ["include", "same-origin", "omit"]) {
      try {
        const r = await fetch(url, { credentials: creds, cache: "force-cache" });
        if (r.ok) {
          const buf = await r.arrayBuffer();
          if (buf && buf.byteLength > 0) return buf;
        }
      } catch { }
    }
    return null;
  }

  const LOGOS_READY = (async function preloadImages() {
    await Promise.all(ALL_LOGO_URLS.map(async url => {
      const buf = await fetchImageBuffer(url);
      if (buf) LOGO_CACHE.set(url, buf);
    }));
    console.log("✅ Logo images preloaded:", LOGO_CACHE.size, "of", ALL_LOGO_URLS.length);
    if (LOGO_CACHE.size < ALL_LOGO_URLS.length) {
      console.warn("⚠️ Some logos failed to load:", ALL_LOGO_URLS.filter(u => !LOGO_CACHE.has(u)));
    }
  })();

  const PAGE_WIDTH_TWIPS = 12240, MARGIN_L = 720, MARGIN_R = 720, CONTENT_W = PAGE_WIDTH_TWIPS - MARGIN_L - MARGIN_R;
  const HEADER_BG = "e7f1e7";

  const LABELS = {
    EN: {
      invoiceNum: "Invoice #",
      invoiceDate: "Invoice Date",
      lcNumber: "L/C Number",
      lcIssueDate: "L/C Issuing Date",
      beneficiary: "Beneficiary / Exporter",
      consignee: "Consignee",
      applicant: "Applicant",
      shipTo: "Ship To",
      billTo: "Bill To",
      shippingMarks: "Shipping Marks/ Brands",
      typeOfTransport: "Type of Transportation",
      transportMode: "Transportation Mode",
      portOfLoading: "Port of Loading",
      portOfDischarge: "Port of Discharge",
      piNumber: "PI Number",
      customerPoNumber: "Customer PO Number",
      incoterms: "Incoterms",
      countryOfOrigin: "Country of Origin",
      currency: "Currency",
      paymentTerms: "Payment Terms",
      sn: "S/N",
      orderNum: "Order #",
      hsCode: "HS Code",
      itemCode: "Item Code",
      itemDesc: "Item Description",
      qty: "Qty.",
      unitPrice: "Unit Price",
      notifyParty: "Notify Party",
      orderNumbers: "Order No.",
      vat5: "5% VAT",
      totalPrice: "Total Price",
      packaging: "Packaging",
      uom: "UOM",
      totalLiters: "Total Liters",
      netWeight: "Net Weight (Kgs)",
      grossWeight: "Gross Weight (Kgs)",
      container: "Container #",
      additionalComments: "Additional Comments / Conditions:",
      termsConditions: "Terms and Conditions #",
      totalPackage: "Total Package",
      subtotal: "Subtotal",
      freightCharges: "Freight Charges",
      insuranceCharges: "Insurance Charges",
      specialDiscount: "Special Discount",
      vat5Total: "5 % VAT Total",
      grandTotal: "Grand Total",
      totalExVat: "Total (Ex VAT)",  // ✅ ADD THIS

      additionalDetails: "Additional Details",
      value: "Value",
      additionalCommentsLabel: "Additional Comments",
      descriptionGoods: "Description of Goods and/or Services"
    },
    AR: {
      invoiceNum: "رقم الفاتورة",
      invoiceDate: "تاريخ الفاتورة",
      lcNumber: "رقم الاعتماد المستندي",
      lcIssueDate: "تاريخ إصدار الاعتماد",
      beneficiary: "المستفيد / المصدر",
      consignee: "المرسل إليه",
      applicant: "مقدم الطلب",
      shipTo: "الشحن إلى",
      billTo: "إرسال الفاتورة إلى",
      shippingMarks: "علامات الشحن / العلامات التجارية",
      typeOfTransport: "نوع النقل",
      transportMode: "وسيلة النقل",
      portOfLoading: "ميناء التحميل",
      portOfDischarge: "ميناء التفريغ",
      piNumber: "رقم الفاتورة الأولية",
      customerPoNumber: "رقم أمر شراء العميل",
      incoterms: "شروط التسليم",
      notifyParty: "الجهة المبلغة",
      orderNumbers: "رقم الطلب",
      totalExVat: "المجموع (بدون ضريبة)",  // ✅ ADD THIS

      countryOfOrigin: "بلد المنشأ",
      currency: "العملة",
      paymentTerms: "شروط الدفع",
      sn: "الرقم التسلسلي",
      orderNum: "رقم الطلب",
      hsCode: "رمز HS",
      itemCode: "رمز الصنف",
      itemDesc: "وصف الصنف",
      qty: "الكمية",
      unitPrice: "سعر الوحدة",
      vat5: "ضريبة القيمة المضافة 5%",
      totalPrice: "السعر الإجمالي",
      packaging: "التعبئة",
      uom: "وحدة القياس",
      totalLiters: "إجمالي اللترات",
      netWeight: "الوزن الصافي (كجم)",
      grossWeight: "الوزن الإجمالي (كجم)",
      container: "حاوية رقم",
      additionalComments: "تعليقات / شروط إضافية:",
      termsConditions: "الشروط والأحكام رقم",
      totalPackage: "إجمالي الحزم",
      subtotal: "المجموع الفرعي",
      freightCharges: "رسوم الشحن",
      insuranceCharges: "رسوم التأمين",
      specialDiscount: "خصم خاص",
      vat5Total: "إجمالي ضريبة القيمة المضافة 5%",
      grandTotal: "المجموع الإجمالي",
      additionalDetails: "تفاصيل إضافية",
      value: "القيمة",
      additionalCommentsLabel: "تعليقات إضافية",
      descriptionGoods: "وصف البضائع و/أو الخدمات"
    }
  };

  // Add after LABELS constant (~line 180)
  function isRTL(lang) {
    return (lang || "").toUpperCase() === "AR";
  }

  function getContainerTypeLabel(value) {
    const typeMap = {
      1: "20 ft.",
      2: "40 ft.",
      3: "Trucks",
      4: "Bulk Tanker"
    };

    // Handle both number and string values
    const numValue = typeof value === 'string' ? parseInt(value, 10) : value;
    return typeMap[numValue] || "";
  }
  /* =========================
   COUNTRY CODE MAPPING
   ========================= */
  const COUNTRY_CODES = {
    "AFGHANISTAN": "AFG", "ALBANIA": "ALB", "ALGERIA": "DZA", "AMERICAN SAMOA": "ASM",
    "ANDORRA": "AND", "ANGOLA": "AGO", "ANTIGUA AND BARBUDA": "ATG", "ARGENTINA": "ARG",
    "ARMENIA": "ARM", "AUSTRALIA": "AUS", "AUSTRIA": "AUT", "AZERBAIJAN": "AZE",
    "BAHAMAS": "BHS", "BAHRAIN": "BHR", "BANGLADESH": "BGD", "BARBADOS": "BRB",
    "BELARUS": "BLR", "BELGIUM": "BEL", "BELIZE": "BLZ", "BENIN": "BEN",
    "BHUTAN": "BTN", "BOLIVIA": "BOL", "BOSNIA AND HERZEGOVINA": "BIH", "BOTSWANA": "BWA",
    "BRAZIL": "BRA", "BRUNEI": "BRN", "BULGARIA": "BGR", "BURKINA FASO": "BFA",
    "BURUNDI": "BDI", "CAMBODIA": "KHM", "CAMEROON": "CMR", "CANADA": "CAN",
    "CAPE VERDE": "CPV", "CENTRAL AFRICAN REPUBLIC": "CAF", "CHAD": "TCD", "CHILE": "CHL",
    "CHINA": "CHN", "COLOMBIA": "COL", "COMOROS": "COM", "CONGO": "COG",
    "COSTA RICA": "CRI", "CROATIA": "HRV", "CUBA": "CUB", "CYPRUS": "CYP",
    "CZECH REPUBLIC": "CZE", "DENMARK": "DNK", "DJIBOUTI": "DJI", "DOMINICA": "DMA",
    "DOMINICAN REPUBLIC": "DOM", "ECUADOR": "ECU", "EGYPT": "EGY", "EL SALVADOR": "SLV",
    "EQUATORIAL GUINEA": "GNQ", "ERITREA": "ERI", "ESTONIA": "EST", "ETHIOPIA": "ETH",
    "FIJI": "FJI", "FINLAND": "FIN", "FRANCE": "FRA", "GABON": "GAB",
    "GAMBIA": "GMB", "GEORGIA": "GEO", "GERMANY": "DEU", "GHANA": "GHA",
    "GREECE": "GRC", "GRENADA": "GRD", "GUATEMALA": "GTM", "GUINEA": "GIN",
    "GUINEA-BISSAU": "GNB", "GUYANA": "GUY", "HAITI": "HTI", "HONDURAS": "HND",
    "HUNGARY": "HUN", "ICELAND": "ISL", "INDIA": "IND", "INDONESIA": "IDN",
    "IRAN": "IRN", "IRAQ": "IRQ", "IRELAND": "IRL", "ISRAEL": "ISR",
    "ITALY": "ITA", "JAMAICA": "JAM", "JAPAN": "JPN", "JORDAN": "JOR",
    "KAZAKHSTAN": "KAZ", "KENYA": "KEN", "KIRIBATI": "KIR", "KOREA, NORTH": "PRK",
    "KOREA, SOUTH": "KOR", "KUWAIT": "KWT", "KYRGYZSTAN": "KGZ", "LAOS": "LAO",
    "LATVIA": "LVA", "LEBANON": "LBN", "LESOTHO": "LSO", "LIBERIA": "LBR",
    "LIBYA": "LBY", "LIECHTENSTEIN": "LIE", "LITHUANIA": "LTU", "LUXEMBOURG": "LUX",
    "MADAGASCAR": "MDG", "MALAWI": "MWI", "MALAYSIA": "MYS", "MALDIVES": "MDV",
    "MALI": "MLI", "MALTA": "MLT", "MARSHALL ISLANDS": "MHL", "MAURITANIA": "MRT",
    "MAURITIUS": "MUS", "MEXICO": "MEX", "MICRONESIA": "FSM", "MOLDOVA": "MDA",
    "MONACO": "MCO", "MONGOLIA": "MNG", "MONTENEGRO": "MNE", "MOROCCO": "MAR",
    "MOZAMBIQUE": "MOZ", "MYANMAR": "MMR", "NAMIBIA": "NAM", "NAURU": "NRU",
    "NEPAL": "NPL", "NETHERLANDS": "NLD", "NEW ZEALAND": "NZL", "NICARAGUA": "NIC",
    "NIGER": "NER", "NIGERIA": "NGA", "NORTH MACEDONIA": "MKD", "NORWAY": "NOR",
    "OMAN": "OMN", "PAKISTAN": "PAK", "PALAU": "PLW", "PANAMA": "PAN",
    "PAPUA NEW GUINEA": "PNG", "PARAGUAY": "PRY", "PERU": "PER", "PHILIPPINES": "PHL",
    "POLAND": "POL", "PORTUGAL": "PRT", "QATAR": "QAT", "ROMANIA": "ROU",
    "RUSSIA": "RUS", "RWANDA": "RWA", "SAINT KITTS AND NEVIS": "KNA", "SAINT LUCIA": "LCA",
    "SAINT VINCENT AND THE GRENADINES": "VCT", "SAMOA": "WSM", "SAN MARINO": "SMR",
    "SAO TOME AND PRINCIPE": "STP", "SAUDI ARABIA": "SAU", "SENEGAL": "SEN",
    "SERBIA": "SRB", "SEYCHELLES": "SYC", "SIERRA LEONE": "SLE", "SINGAPORE": "SGP",
    "SLOVAKIA": "SVK", "SLOVENIA": "SVN", "SOLOMON ISLANDS": "SLB", "SOMALIA": "SOM",
    "SOUTH AFRICA": "ZAF", "SOUTH SUDAN": "SSD", "SPAIN": "ESP", "SRI LANKA": "LKA",
    "SUDAN": "SDN", "SURINAME": "SUR", "SWEDEN": "SWE", "SWITZERLAND": "CHE",
    "SYRIA": "SYR", "TAIWAN": "TWN", "TAJIKISTAN": "TJK", "TANZANIA": "TZA",
    "THAILAND": "THA", "TIMOR-LESTE": "TLS", "TOGO": "TGO", "TONGA": "TON",
    "TRINIDAD AND TOBAGO": "TTO", "TUNISIA": "TUN", "TURKEY": "TUR", "TURKMENISTAN": "TKM",
    "TUVALU": "TUV", "UGANDA": "UGA", "UKRAINE": "UKR", "UNITED ARAB EMIRATES": "ARE",
    "UAE": "ARE", "UNITED KINGDOM": "GBR", "UK": "GBR", "UNITED STATES": "USA",
    "USA": "USA", "URUGUAY": "URY", "UZBEKISTAN": "UZB", "VANUATU": "VUT",
    "VATICAN CITY": "VAT", "VENEZUELA": "VEN", "VIETNAM": "VNM", "YEMEN": "YEM",
    "ZAMBIA": "ZMB", "ZIMBABWE": "ZWE"
  };

  /**
   * Get country code from country name
   */
  function getCountryCode(countryName) {
    if (!countryName) return "UNK";

    const normalized = countryName.trim().toUpperCase();

    // Direct match
    if (COUNTRY_CODES[normalized]) {
      return COUNTRY_CODES[normalized];
    }

    // Partial match
    for (const [key, code] of Object.entries(COUNTRY_CODES)) {
      if (normalized.includes(key) || key.includes(normalized)) {
        return code;
      }
    }

    return "UNK";
  }

  /**
   * Generate CI Number from brand, auto-number, and country
   * Format: {BrandInitial}-{FirstPart}-{CountryCode}-{SecondPart}
   * Example: P-1125-IRQ-0000
   */
  function generateCINumber(firstBrand, autoNumberIC, countryName) {
    console.log("🔧 generateCINumber called with:", { firstBrand, autoNumberIC, countryName });

    if (!firstBrand || !autoNumberIC) {
      console.warn("❌ Cannot generate CI number: missing brand or auto-number");
      console.warn("   - firstBrand:", firstBrand);
      console.warn("   - autoNumberIC:", autoNumberIC);
      return null;
    }

    // Get first letter of brand (uppercase)
    const brandInitial = firstBrand.trim().charAt(0).toUpperCase();
    console.log("   - Brand Initial:", brandInitial);

    // Split auto-number (e.g., "1125-0000" -> ["1125", "0000"])
    const parts = autoNumberIC.split("-");
    console.log("   - Auto-number parts:", parts);

    if (parts.length !== 2) {
      console.warn("❌ Invalid auto-number format:", autoNumberIC);
      return null;
    }

    const firstPart = parts[0].trim();
    const secondPart = parts[1].trim();
    console.log("   - First part:", firstPart);
    console.log("   - Second part:", secondPart);

    // Get country code
    const countryCode = getCountryCode(countryName);
    console.log("   - Country code:", countryCode);

    // Combine: BrandInitial-FirstPart-CountryCode-SecondPart
    const result = `${brandInitial}-${firstPart}-${countryCode}-${secondPart}`;
    console.log("✅ Generated CI Number:", result);

    return result;
  }

  /**
   * Update CI Number in master record
   */
  async function updateCINumber(dclId, ciNumber) {
    if (!dclId || !ciNumber) {
      console.warn("Cannot update CI number: missing ID or number");
      return;
    }

    const url = `${API_URL_MASTER}(${dclId})`;

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

    const headers = {
      "Accept": "application/json;odata.metadata=minimal",
      "Content-Type": "application/json; charset=utf-8",
      "If-Match": "*"
    };

    const token = await getAntiForgeryTokenMaybe();
    if (token) {
      headers["__RequestVerificationToken"] = token;
    }

    const body = {
      cr650_ci_number: ciNumber
    };

    Loading.start();
    try {
      const res = await fetch(url, {
        method: "PATCH",
        headers,
        body: JSON.stringify(body)
      });

      if (!res.ok) {
        const txt = await res.text();
        console.error("Failed to update CI number", res.status, txt);
      } else {
        console.log("✅ CI number updated successfully:", ciNumber);
      }
    } catch (err) {
      console.error("Error updating CI number:", err);
    } finally {
      Loading.stop();
    }
  }

  /* =========================
     3) UTILITIES (shared)
     ========================= */
  function getQueryId() { const m = /[?&]id=([^&#]+)/i.exec(location.href); return m ? decodeURIComponent(m[1]) : null; }
  function guidOK(g) { return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(g || ""); }
  function fmtNum(n, dec = 2) { const v = Number(String(n ?? "").toString().replace(/,/g, "")); if (isNaN(v)) return ""; return v.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec }); }
  function n(v, dec = 0) { const x = Number(v); return Number.isFinite(x) ? x.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec }) : ""; }
  async function safeAjax(url) { Loading.start(); try { const r = await fetch(url, { headers: { Accept: "application/json" } }); const t = await r.text(); if (!r.ok) throw new Error(t || r.statusText); try { return JSON.parse(t); } catch { return t; } } finally { Loading.stop(); } }
  async function ensureDocx() { if (w.docx?.Packer) return; Loading.start(); try { await new Promise((res, rej) => { const s = d.createElement("script"); s.src = "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js"; s.onload = res; s.onerror = () => rej(new Error("docx load failed")); d.head.appendChild(s); }); } finally { Loading.stop(); } }
  async function fetchArrayBuffer(url) {
    await LOGOS_READY;
    if (LOGO_CACHE.has(url)) return LOGO_CACHE.get(url);
    const buf = await fetchImageBuffer(url);
    if (buf) LOGO_CACHE.set(url, buf);
    return buf;
  }
  function base64LenToBytes(len) { return Math.floor((len * 3) / 4); }
  async function blobToBase64Raw(blob) { const buf = await blob.arrayBuffer(), bytes = new Uint8Array(buf); let b = ""; const chunk = 0x8000; for (let i = 0; i < bytes.length; i += chunk) b += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)); return btoa(b); }

  function qp(obj) { return Object.entries(obj).map(([k, v]) => k + "=" + encodeURIComponent(v)).join("&"); }
  function buildHrefWithId(href, id) {
    if (!id) return href;
    try { const u = new URL(href, location.origin); u.searchParams.set("id", id); const sp = u.searchParams.toString(); return u.pathname + (sp ? "?" + sp : "") + (u.hash || ""); }
    catch { const sep = href.includes("?") ? "&" : "?"; return href + sep + "id=" + encodeURIComponent(id); }
  }
  function rewriteWizardLinksWithId(id) {
    if (!id) return;
    const $scope = $("#dclWizard").length ? $("#dclWizard") : $(d);
    $scope.find("#stepIndicators a, .navigation-buttons a").each(function () {
      const raw = $(this).attr("href"); if (!raw) return;
      $(this).attr("href", buildHrefWithId(raw, id));
    });
  }
  function wireNavLoaderClicks() {
    const $scope = $("#dclWizard").length ? $("#dclWizard") : $(d);
    $scope.find("#stepIndicators a, .navigation-buttons a").each(function () {
      $(this).on("click", function () { Loading.start(); });
    });
  }

  function getCompanyType(masterRec) {
    const val = masterRec[MF.companyType];
    if (val === 0 || val === "0") return "TECHNOLUBE";
    if (val === 1 || val === "1") return "PETROLUBE";
    return "PETROLUBE";
  }
  /* =========================
   * FORM LOCKING FOR SUBMITTED DCLs
   * ========================= */
  function lockFormIfSubmitted() {
    try {
      const status = (DCL_STATUS || "").toLowerCase();

      if (status !== "submitted") {
        console.log("📝 Form is editable - status:", DCL_STATUS || "none");
        return;
      }

      console.log("🔒 Locking document generator - DCL status is 'Submitted'");

      // 1. Disable all radio buttons (language selection)
      document.querySelectorAll('input[type="radio"][name^="ci_lang"], input[type="radio"][name^="pl_lang"]').forEach(radio => {
        radio.disabled = true;
        radio.style.cursor = "not-allowed";
        const label = radio.closest("label");
        if (label) {
          label.style.opacity = "0.6";
          label.style.cursor = "not-allowed";
        }
      });

      // 2. Disable and hide all Generate buttons
      const generateButtons = document.querySelectorAll("#btnGenCI, #btnGenPL");
      generateButtons.forEach(btn => {
        btn.disabled = true;
        btn.style.display = "none";
      });

      // 3. Disable and hide all Regenerate buttons
      const regenButtons = document.querySelectorAll('[id^="btnRegen"]');
      regenButtons.forEach(btn => {
        btn.disabled = true;
        btn.style.display = "none";
      });

      // 4. Make View buttons read-only (keep visible but indicate locked state)
      const viewButtons = document.querySelectorAll("#btnViewCI, #btnPrevPL");
      viewButtons.forEach(btn => {
        // Add a visual indicator that the document is locked
        const icon = btn.querySelector("i");
        if (icon) {
          icon.classList.remove("fa-link");
          icon.classList.add("fa-lock");
        }
      });

      // 5. Disable document cards interaction
      document.querySelectorAll(".doc-card").forEach(card => {
        card.style.opacity = "0.8";
        card.style.pointerEvents = "none";

        // Re-enable view links only
        const viewLink = card.querySelector("a.btn-outline");
        if (viewLink && viewLink.href && viewLink.href !== "#") {
          viewLink.style.pointerEvents = "auto";
          viewLink.style.opacity = "1";
        }
      });

      // 6. Lock discounts/charges section
      const discountsSection = document.querySelector(".discounts-charges-section");
      if (discountsSection) {
        discountsSection.style.opacity = "0.8";
        discountsSection.style.pointerEvents = "none";

        // Disable all inputs and buttons in discounts section
        discountsSection.querySelectorAll("input, select, button").forEach(el => {
          el.disabled = true;
          el.style.cursor = "not-allowed";
        });
      }

      // 7. Show locked banner
      showLockedBanner();

      console.log("✅ Document generator fully locked");

    } catch (lockError) {
      console.error("❌ Error while locking document generator:", lockError);
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
          This DCL has been submitted. You can view existing documents but cannot generate new ones.
        </div>
      </div>
    `;

      // Insert before the document grid
      const docGrid = wizard.querySelector(".doc-grid");
      if (docGrid) {
        wizard.insertBefore(banner, docGrid);
      } else {
        wizard.insertBefore(banner, wizard.firstChild);
      }
    } catch (bannerError) {
      console.error("Error creating locked banner:", bannerError);
    }
  }

  /* =============================
     DISCOUNTS / CHARGES MODULE
     ============================= */
  const discountChargeRecords = new Map(); // localId -> dataverse GUID
  w.DISCOUNT_CHARGE_NET = 0;

  // Helper function to make AJAX calls with CSRF token for write operations
  function safeAjaxWithToken(options) {
    return new Promise((resolve, reject) => {
      function executeRequest(token) {
        if (token) {
          options.headers = options.headers || {};
          options.headers["__RequestVerificationToken"] = token;
        }

        $.ajax(options)
          .done((data, textStatus, jqXHR) => {
            // Handle validateLoginSession if available
            if (typeof w.validateLoginSession === "function") {
              try {
                w.validateLoginSession(data, textStatus, jqXHR, resolve);
              } catch {
                resolve(data);
              }
            } else {
              resolve(data);
            }
          })
          .fail((jqXHR, textStatus, errorThrown) => {
            reject(jqXHR);
          });
      }

      // Get CSRF token for write operations
      if (w.shell && typeof w.shell.getTokenDeferred === "function") {
        w.shell.getTokenDeferred()
          .done(executeRequest)
          .fail(() => reject({ message: "Token API unavailable" }));
      } else {
        // Fallback: try without token (may fail for write operations)
        executeRequest(null);
      }
    });
  }

  function createTypeInput(selectedValue = "") {
    const options = [
      "Marketing Materials",
      "Special Discount",
      "Other Charges",
      "Less Advance Payment Received",
      "Insurance",
      "Freight Charges",
      "Flexi Bags",
      "Documentation Charges"
    ];

    let html = `<select class="dc-type">`;
    html += `<option value="">Select...</option>`;

    for (const opt of options) {
      const sel = opt === selectedValue ? "selected" : "";
      html += `<option value="${opt}" ${sel}>${opt}</option>`;
    }

    html += `</select>`;
    return html;
  }

  function addDiscountChargeRow() {
    const tbody = d.querySelector("#discountsChargesTable tbody");
    if (!tbody) return;

    // Remove empty-row if it exists
    const emptyRow = tbody.querySelector(".empty-row");
    if (emptyRow) {
      emptyRow.remove();
    }

    const tr = d.createElement("tr");

    const localId = w.crypto && w.crypto.randomUUID
      ? w.crypto.randomUUID()
      : String(Date.now() + Math.random());

    tr.dataset.recordId = localId;
    tr.dataset.isNew = "true";

    tr.innerHTML = `
      <td>
        <select class="dc-type">
          <option value="">Select...</option>
          <option value="Marketing Materials">Marketing Materials</option>
          <option value="Special Discount">Special Discount</option>
          <option value="Other Charges">Other Charges</option>
          <option value="Less Advance Payment Received">Less Advance Payment Received</option>
          <option value="Insurance">Insurance</option>
          <option value="Freight Charges">Freight Charges</option>
          <option value="Flexi Bags">Flexi Bags</option>
          <option value="Documentation Charges">Documentation Charges</option>
        </select>
        <input type="text" class="dc-other-name" placeholder="Enter charge name...">
      </td>
      <td>
        <input type="number" class="dc-quantity" value="1" min="1" step="1">
      </td>
      <td>
        <input type="number" class="dc-amount" value="0" step="0.01">
      </td>
      <td>
        <select class="dc-currency">
          <option value="USD">USD</option>
          <option value="SAR" selected>SAR</option>
          <option value="AED">AED</option>
        </select>
      </td>
      <td>
        <button type="button" class="btn btn-danger btn-sm dc-remove">
          <i class="fas fa-trash"></i> Remove
        </button>
      </td>
    `;

    tbody.appendChild(tr);

    const removeBtn = tr.querySelector(".dc-remove");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        if (w.confirm("Remove this item?")) {
          tr.remove();

          // Show empty-row again if no rows left
          const remainingRows = tbody.querySelectorAll("tr:not(.empty-row)");
          if (remainingRows.length === 0) {
            tbody.innerHTML = `
            <tr class="empty-row">
              <td colspan="5">
                <i class="fas fa-inbox"></i>
                No discounts or charges added yet
              </td>
            </tr>
          `;
          }

          updateDiscountChargeSummary();
        }
      });
    }

    // Show/hide "Other Charges" name field based on type selection
    const typeSelect = tr.querySelector(".dc-type");
    const otherNameInput = tr.querySelector(".dc-other-name");
    typeSelect.addEventListener("change", () => {
      otherNameInput.style.display = typeSelect.value === "Other Charges" ? "block" : "none";
      if (typeSelect.value !== "Other Charges") otherNameInput.value = "";
    });

    // Add event listeners for live summary updates
    tr.querySelector(".dc-amount").addEventListener("input", updateDiscountChargeSummary);
    tr.querySelector(".dc-type").addEventListener("change", updateDiscountChargeSummary);
    tr.querySelector(".dc-quantity").addEventListener("input", updateDiscountChargeSummary);

    updateDiscountChargeSummary();
  }

  async function saveDiscountRowToDataverse(row, dataverseId = null) {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) {
      console.error("No DCL ID found – cannot save discounts/charges");
      throw new Error("No DCL ID");
    }

    const payload = {
      cr650_name: row.type,
      cr650_amount: row.amount,
      cr650_currency: row.currency,
      cr650_qty: row.quantity || 1,
      cr650_totalimpact: row.amount
    };

    if (row.type === "Other Charges") {
      payload.cr650_other_charges_name = row.otherChargesName || "";
    } else {
      payload.cr650_other_charges_name = null;
    }

    if (!dataverseId) {
      payload["cr650_DCLReference@odata.bind"] = `/cr650_dcl_masters(${dclGuid})`;
    }

    if (!dataverseId) {
      // CREATE
      try {
        console.log("🔵 Creating discount/charge:", row.type);

        await safeAjaxWithToken({
          type: "POST",
          url: "/_api/cr650_dcl_discounts_chargeses",
          data: JSON.stringify(payload),
          contentType: "application/json; charset=utf-8",
          headers: { Accept: "application/json" },
          dataType: "json"
        });

        console.log("✅ Record created, now fetching it back...");

        // Fetch the newly created record
        await new Promise(resolve => setTimeout(resolve, 100));

        const escapedType = row.type.replace(/'/g, "''");
        const fetchUrl = `/_api/cr650_dcl_discounts_chargeses?$filter=_cr650_dclreference_value eq '${dclGuid}' and cr650_name eq '${escapedType}' and cr650_amount eq ${row.amount}&$orderby=createdon desc&$top=1`;

        const fetchRes = await $.ajax({
          url: fetchUrl,
          headers: { Accept: "application/json" }
        });

        if (fetchRes && fetchRes.value && fetchRes.value.length > 0) {
          const id = fetchRes.value[0].cr650_dcl_discounts_chargesid;
          console.log("✅ Successfully retrieved ID:", id);
          return id;
        }

        // Alternative fetch
        const altFetchUrl = `/_api/cr650_dcl_discounts_chargeses?$filter=_cr650_dclreference_value eq '${dclGuid}'&$orderby=createdon desc&$top=1`;
        const altFetchRes = await $.ajax({ url: altFetchUrl });

        if (altFetchRes && altFetchRes.value && altFetchRes.value.length > 0) {
          return altFetchRes.value[0].cr650_dcl_discounts_chargesid;
        }

        throw new Error("Could not find newly created discount/charge record");

      } catch (err) {
        console.error("❌ CREATE/FETCH failed:", err);
        throw err;
      }
    }

    // UPDATE
    try {
      console.log("🔄 Updating discount/charge:", dataverseId);

      await safeAjaxWithToken({
        type: "PATCH",
        url: `/_api/cr650_dcl_discounts_chargeses(${dataverseId})`,
        data: JSON.stringify(payload),
        contentType: "application/json; charset=utf-8",
        headers: {
          Accept: "application/json;odata.metadata=minimal",
          "If-Match": "*"
        },
        dataType: "json"
      });

      console.log("✅ UPDATE successful");
      return dataverseId;

    } catch (err) {
      console.error("❌ UPDATE failed:", err);
      throw err;
    }
  }

  function readDiscountChargeRows() {
    const rows = [];
    const trs = d.querySelectorAll("#discountsChargesTable tbody tr") || [];

    trs.forEach(tr => {
      if (tr.classList.contains("empty-row")) return;

      const type = tr.querySelector(".dc-type")?.value?.trim() || "";
      const amountRaw = tr.querySelector(".dc-amount")?.value;
      const currency = tr.querySelector(".dc-currency")?.value || "USD";
      const quantityRaw = tr.querySelector(".dc-quantity")?.value;

      if (!type && (!amountRaw || amountRaw === "0")) return;

      let amount = parseFloat(amountRaw);
      if (isNaN(amount)) amount = 0;

      let quantity = parseInt(quantityRaw);
      if (isNaN(quantity) || quantity < 1) quantity = 1;

      const recordId = tr.dataset.recordId;
      const isNew = tr.dataset.isNew === "true";
      const dataverseId = discountChargeRecords.get(recordId) || null;

      const otherChargesName = type === "Other Charges"
        ? (tr.querySelector(".dc-other-name")?.value?.trim() || "")
        : "";

      rows.push({
        type,
        amount,
        currency,
        quantity,
        recordId,
        isNew,
        dataverseId,
        otherChargesName
      });
    });

    return rows;
  }

  async function saveAllDiscountsAndCharges() {
    const rows = readDiscountChargeRows();

    if (!rows.length) {
      showToast("warning", "No discount/charge rows to save.");
      return;
    }

    Loading.start();

    let savedCount = 0;
    let errorCount = 0;

    try {
      for (const row of rows) {
        try {
          if (row.isNew) {
            const newGuid = await saveDiscountRowToDataverse(row, null);
            const tr = d.querySelector(`tr[data-record-id="${row.recordId}"]`);
            if (tr) tr.dataset.isNew = "false";
            discountChargeRecords.set(row.recordId, newGuid);
          } else if (row.dataverseId) {
            await saveDiscountRowToDataverse(row, row.dataverseId);
          }
          savedCount++;
        } catch (err) {
          console.error(`Failed to save discount/charge "${row.type}"`, err);
          errorCount++;
        }
      }

      if (errorCount > 0) {
        showToast("error", `Saved ${savedCount}, but ${errorCount} item(s) failed.`);
      } else {
        showToast("success", `Successfully saved ${savedCount} discount/charge item(s).`);
      }

      updateDiscountChargeSummary();
    } finally {
      Loading.stop();
    }
  }

  async function loadDiscountsFromDataverse() {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) return [];

    try {
      Loading.start();
      const results = await $.ajax({
        type: "GET",
        url: `/_api/cr650_dcl_discounts_chargeses?$filter=_cr650_dclreference_value eq '${dclGuid}'`,
        headers: { Accept: "application/json;odata.metadata=minimal" },
        dataType: "json"
      });

      const rows = (results.value || []).map(r => ({
        id: r.cr650_dcl_discounts_chargesid,
        type: r.cr650_name,
        amount: r.cr650_amount || 0,
        currency: r.cr650_currency || "USD",
        quantity: r.cr650_qty || 1,
        otherChargesName: r.cr650_other_charges_name || ""
      }));

      return rows;
    } catch (err) {
      console.error("Error loading discounts/charges from Dataverse:", err);
      return [];
    } finally {
      Loading.stop();
    }
  }

  function renderDiscountCharges(rows) {
    const tbody = d.querySelector("#discountsChargesTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";
    discountChargeRecords.clear();

    if (!rows || rows.length === 0) {
      tbody.innerHTML = `
        <tr class="empty-row">
          <td colspan="5">
            <i class="fas fa-inbox"></i>
            No discounts or charges added yet
          </td>
        </tr>
      `;
      updateDiscountChargeSummary();
      return;
    }

    rows.forEach(r => {
      const tr = d.createElement("tr");
      const localId = w.crypto && w.crypto.randomUUID
        ? w.crypto.randomUUID()
        : String(Date.now() + Math.random());

      tr.dataset.recordId = localId;
      tr.dataset.isNew = "false";

      if (r.id) {
        discountChargeRecords.set(localId, r.id);
      }

      const isOtherCharges = r.type === "Other Charges";
      const otherNameVal = r.otherChargesName || "";

      tr.innerHTML = `
        <td class="cell" align="left">
          ${createTypeInput(r.type)}
          <input type="text" class="dc-other-name" placeholder="Enter charge name..." value="${otherNameVal.replace(/"/g, '&quot;')}"${isOtherCharges ? ' style="display:block"' : ''}>
        </td>
        <td class="cell" align="left">
          <input type="number" class="dc-quantity" value="${r.quantity || 1}" min="1" step="1" style="width:100%;">
        </td>
        <td class="cell" align="left">
          <input type="number" class="dc-amount" value="${r.amount}" step="0.01" style="width:100%;">
        </td>
        <td class="cell" align="left">
          <select class="dc-currency">
            <option value="USD" ${r.currency === "USD" ? "selected" : ""}>USD</option>
            <option value="SAR" ${r.currency === "SAR" ? "selected" : ""}>SAR</option>
            <option value="AED" ${r.currency === "AED" ? "selected" : ""}>AED</option>
          </select>
        </td>
        <td class="cell" align="left">
          <button type="button" class="btn btn-danger btn-sm dc-remove">
            <i class="fas fa-trash"></i> Remove
          </button>
        </td>
      `;

      tbody.appendChild(tr);

      // Show/hide "Other Charges" name field based on type selection
      const typeSelect = tr.querySelector(".dc-type");
      const otherNameInput = tr.querySelector(".dc-other-name");
      typeSelect.addEventListener("change", () => {
        otherNameInput.style.display = typeSelect.value === "Other Charges" ? "block" : "none";
        if (typeSelect.value !== "Other Charges") otherNameInput.value = "";
      });

      tr.querySelector(".dc-remove").addEventListener("click", async () => {
        if (!w.confirm("Remove this item?")) return;

        const guid = discountChargeRecords.get(localId);
        if (guid) {
          try {
            Loading.start();
            await safeAjaxWithToken({
              type: "DELETE",
              url: `/_api/cr650_dcl_discounts_chargeses(${guid})`,
              headers: {
                Accept: "application/json;odata.metadata=minimal",
                "If-Match": "*"
              },
              dataType: "json"
            });
          } catch (err) {
            console.error("Delete failed:", err);
          } finally {
            Loading.stop();
          }
        }

        discountChargeRecords.delete(localId);
        tr.remove();

        // Show empty-row if no rows left
        const remainingRows = tbody.querySelectorAll("tr:not(.empty-row)");
        if (remainingRows.length === 0) {
          tbody.innerHTML = `
            <tr class="empty-row">
              <td colspan="5">
                <i class="fas fa-inbox"></i>
                No discounts or charges added yet
              </td>
            </tr>
          `;
        }

        updateDiscountChargeSummary();
      });

      tr.querySelector(".dc-amount").addEventListener("input", updateDiscountChargeSummary);
      tr.querySelector(".dc-type").addEventListener("change", updateDiscountChargeSummary);
      tr.querySelector(".dc-quantity").addEventListener("input", updateDiscountChargeSummary);
    });

    updateDiscountChargeSummary();
  }

  function updateDiscountChargeSummary() {
    let totalCharges = 0;
    let totalDiscounts = 0;

    const trs = d.querySelectorAll("#discountsChargesTable tbody tr") || [];

    trs.forEach(tr => {
      if (tr.classList.contains("empty-row")) return;

      const type = tr.querySelector(".dc-type")?.value || "";
      const rawAmount = tr.querySelector(".dc-amount")?.value;
      const amount = parseFloat(rawAmount) || 0;

      if (!type) return;

      // Use the sign of the amount: Negative = Discount, Positive = Additional Charge
      if (amount < 0) {
        totalDiscounts += Math.abs(amount);
      } else {
        totalCharges += amount;
      }
    });

    const netImpact = totalCharges - totalDiscounts;
    w.DISCOUNT_CHARGE_NET = netImpact;

    const summaryDiv = d.querySelector(".dc-summary");
    if (summaryDiv) {
      const hasRows = trs.length > 0 && !trs[0].classList.contains("empty-row");
      summaryDiv.style.display = hasRows ? "block" : "none";

      const chargesEl = summaryDiv.querySelector(".total-charges");
      const discountsEl = summaryDiv.querySelector(".total-discounts");
      const netEl = summaryDiv.querySelector(".net-impact");

      if (chargesEl) chargesEl.textContent = `+${totalCharges.toFixed(2)}`;
      if (discountsEl) discountsEl.textContent = `-${totalDiscounts.toFixed(2)}`;
      if (netEl) netEl.textContent = `${netImpact >= 0 ? "+" : ""}${netImpact.toFixed(2)}`;
    }
  }

  async function initDiscountCharges() {
    const addDCBtn = d.querySelector("#addDiscountChargeBtn");
    if (addDCBtn && !addDCBtn.dataset.bound) {
      addDCBtn.addEventListener("click", addDiscountChargeRow);
      addDCBtn.dataset.bound = "true";
    }

    const saveDCBtn = d.querySelector("#saveDiscountsBtn");
    if (saveDCBtn && !saveDCBtn.dataset.bound) {
      saveDCBtn.addEventListener("click", saveAllDiscountsAndCharges);
      saveDCBtn.dataset.bound = "true";
    }

    const loadDCBtn = d.querySelector("#loadDiscountsBtn");
    if (loadDCBtn && !loadDCBtn.dataset.bound) {
      loadDCBtn.addEventListener("click", async () => {
        const rows = await loadDiscountsFromDataverse();
        renderDiscountCharges(rows);
        showToast("success", "Discounts/Charges refreshed.");
      });
      loadDCBtn.dataset.bound = "true";
    }

    const existingRows = await loadDiscountsFromDataverse();
    if (existingRows.length > 0) {
      renderDiscountCharges(existingRows);
    } else {
      renderDiscountCharges([]);
    }
  }

  /* =============================
     ORDER ITEMS MODULE (Full Editable - Same as Loading Plan View)
     ============================= */

  // Global state
  let CURRENT_DCL_ID = null;
  let CURRENCY_CODE = "USD";
  let ADDITIONAL_DETAILS_ID = null;
  const hiddenDetails = new Set();

  // Helper functions
  const Q = sel => d.querySelector(sel);
  const QA = sel => d.querySelectorAll(sel);
  const asNum = v => parseFloat(String(v || "0").replace(/[^\d.-]/g, "")) || 0;
  const fmt2 = n => fmtNum(n, 2);
  const setText = (sel, val) => { const el = Q(sel); if (el) el.textContent = val; };
  const escapeHtml = s => String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  // Parse UOM from packaging description (e.g., "30x4" = 120, "20L" = 20)
  function parsePackLiters(packDesc) {
    if (!packDesc) return 0;
    const s = String(packDesc).replace(/\s+/g, "").toUpperCase();
    // Pattern 1: "5x4L" or "5X4L" (with L) → 5 * 4 = 20
    const mMultiWithL = s.match(/(\d+)[X×](\d+(\.\d+)?)L/);
    if (mMultiWithL) return Number(mMultiWithL[1]) * Number(mMultiWithL[2]);
    // Pattern 2: "5x4" or "5X4" (without L) → 5 * 4 = 20
    const mMultiNoL = s.match(/^(\d+)[X×](\d+(\.\d+)?)$/);
    if (mMultiNoL) return Number(mMultiNoL[1]) * Number(mMultiNoL[2]);
    // Pattern 3: "20L" or "20 LITER" → 20
    const mSingle = s.match(/(\d+(\.\d+)?)L/);
    if (mSingle) return Number(mSingle[1]);
    // Pattern 4: Just a number "208" → 208
    const mNum = s.match(/(\d+)/);
    return mNum ? Number(mNum[1]) : 0;
  }

  // Detail labels for dropdown
  const DETAIL_LABELS = {
    totalPallets: "Number of pallets the products are packed on",
    totalCartons: "Total Number of Cartons",
    totalDrums: "Total Number of Drums",
    totalPails: "Total Number of Pails",
    totalPackages: "Total Number of Packages",
    totalNetWeight: "Total Net Weight of Packages (Kgs)",
    totalGrossWeight: "Total Gross Weight of Packages (Kgs)"
  };

  // Extended field mapping for loading plan items (aligned with loading_plan_view)
  const LP_FIELDS = {
    id: "cr650_dcl_loading_planid",
    orderNo: "cr650_ordernumber",
    itemCode: "cr650_itemcode",
    hsCode: "cr650_hscode",
    description: "cr650_itemdescription",
    releaseStatus: "cr650_releasestatus",
    packaging: "cr650_packagingdetails",
    uom: "cr650_unitofmeasure",
    pack: "cr650_packagetype",
    orderQty: "cr650_orderedquantity",
    loadingQty: "cr650_loadedquantity",
    pendingQty: "cr650_pendingquantity",
    palletized: "cr650_ispalletized",
    pallets: "cr650_palletcount",
    palletsWeight: "cr650_palletweight",
    totalLiters: "cr650_totalvolumeorweight",
    netKg: "cr650_netweightkg",
    grossKg: "cr650_grossweightkg",
    valueType: "cr650_valuetype",
    unitPrice: "cr650_unitprice",
    vat: "cr650_vat5percent",
    totalExVat: "cr650_totalexcludingvat",
    totalIncVat: "cr650_totalincludingvat",
    containerNo: "cr650_containernumber"
  };

  // Parse loading plan row from Dataverse response (aligned with loading_plan_view)
  function parseLpRow(lpRow) {
    // Convert boolean cr650_ispalletized to "Yes"/"No" string for UI
    const palletizedBool = lpRow.cr650_ispalletized;
    const palletizedStr = (palletizedBool === true || palletizedBool === "Yes") ? "Yes" : "No";

    return {
      serverId: lpRow.cr650_dcl_loading_planid || null,
      orderNo: lpRow.cr650_ordernumber || "",
      itemCode: lpRow.cr650_itemcode || "",
      description: lpRow.cr650_itemdescription || "",
      releaseStatus: lpRow.cr650_releasestatus || "",
      packaging: lpRow.cr650_packagingdetails || "",
      uom: asNum(lpRow.cr650_unitofmeasure),
      pack: lpRow.cr650_packagetype || "",
      OrderQuantity: asNum(lpRow.cr650_orderedquantity),
      LoadingQuantity: asNum(lpRow.cr650_loadedquantity),
      PendingQuantity: asNum(lpRow.cr650_pendingquantity),
      palletized: palletizedStr,
      numberOfPallets: asNum(lpRow.cr650_palletcount),
      palletsWeight: asNum(lpRow.cr650_palletweight),
      totalLiters: asNum(lpRow.cr650_totalvolumeorweight),
      netWeight: asNum(lpRow.cr650_netweightkg),
      grossWeight: asNum(lpRow.cr650_grossweightkg),
      valueType: lpRow.cr650_valuetype || "Sale Price",
      unitPrice: asNum(lpRow.cr650_unitprice),
      vatAmount: asNum(lpRow.cr650_vat5percent),
      totalExVat: asNum(lpRow.cr650_totalexcludingvat),
      totalInclVat: asNum(lpRow.cr650_totalincludingvat),
      containerNumber: lpRow.cr650_containernumber || "",
      hsCode: lpRow.cr650_hscode || "00000000"
    };
  }

  // Create table row element
  function makeRowEl(item, index) {
    const tr = d.createElement("tr");
    tr.classList.add("lp-data-row");
    if (item.serverId) tr.dataset.serverId = item.serverId;

    tr.innerHTML = `
    <td class="sn lp-locked">${index + 1}</td>
    <td class="order-no lp-locked">${escapeHtml(item.orderNo)}</td>
    <td class="item-code lp-locked">${escapeHtml(item.itemCode)}</td>
    <td class="description lp-locked">${escapeHtml(item.description)}</td>
    <td class="release-status-cell lp-locked">
      <select class="release-status-select" disabled>
        <option value="Y" ${item.releaseStatus === "Y" ? "selected" : ""}>Y</option>
        <option value="N" ${item.releaseStatus === "N" ? "selected" : ""}>N</option>
      </select>
    </td>
    <td class="packaging lp-locked">${escapeHtml(item.packaging)}</td>
    <td class="uom ce-num lp-locked">${fmt2(item.uom)}</td>
    <td class="pack lp-locked">${escapeHtml(item.pack)}</td>
    <td class="order-qty ce-num lp-locked">${fmt2(item.OrderQuantity)}</td>
    <td class="loading-qty-cell lp-locked">
      <input type="number" class="loading-qty" min="0" value="${fmt2(item.LoadingQuantity)}" style="width:100px" disabled>
    </td>
    <td class="pending-qty ce-num lp-locked">${fmt2(item.PendingQuantity)}</td>
    <td class="palletized-cell lp-locked">
      <select class="palletized-select" disabled>
        <option value="No" ${item.palletized === "No" ? "selected" : ""}>No</option>
        <option value="Yes" ${item.palletized === "Yes" ? "selected" : ""}>Yes</option>
      </select>
    </td>
    <td class="pallets ce-num lp-locked">${fmt2(item.numberOfPallets)}</td>
    <td class="pallets-weight ce-num lp-locked">${fmt2(item.palletsWeight)}</td>
    <td class="total-liters ce-num lp-locked">${fmt2(item.totalLiters)}</td>
    <td class="net-weight ce-num lp-locked">${fmt2(item.netWeight)}</td>
    <td class="gross-weight ce-num lp-locked">${fmt2(item.grossWeight)}</td>
    <td class="value-type-cell">
      <select class="value-type">
        <option value="Sale Price" ${item.valueType === "Sale Price" ? "selected" : ""}>Sale Price</option>
        <option value="FOC value" ${item.valueType === "FOC value" ? "selected" : ""}>FOC value</option>
        <option value="Priceless" ${item.valueType === "Priceless" ? "selected" : ""}>Priceless</option>
      </select>
    </td>
    <td class="unit-price-cell">
      <input type="number" class="unit-price" min="0" step="any" value="${fmt2(item.unitPrice)}" style="width:110px">
    </td>
    <td class="vat ce-num ce-editable" contenteditable="true">${fmt2(item.vatAmount)}</td>
    <td class="total-ex ce-num ce-editable" contenteditable="true">${fmt2(item.totalExVat)}</td>
    <td class="total-inc ce-num ce-editable" contenteditable="true">${fmt2(item.totalInclVat)}</td>
    <td class="container-no-cell">
      <input type="text" class="container-no" value="${escapeHtml(item.containerNumber || '')}" style="width:200px">
    </td>
    <td class="hs-code-cell">
      <input type="text" class="hs-code" value="${escapeHtml(item.hsCode)}" style="width:120px">
    </td>
    `;
    return tr;
  }

  /**
   * Recalculate derived values for a loading plan row
   * Formulas:
   *   UOM = parse from packing (e.g., 30x4 = 120)
   *   Pending Quantity = Order Quantity − Loading Quantity
   *   Total Liters = Loading Quantity × UOM
   *   Net Weight = Total Liters × DENSITY (1.07 for COOLANT, 0.9 otherwise)
   *   Pallet Weight = 0 if Palletized="No", else Number of Pallets × 19.38
   *   Gross Weight = Pallet Weight + Net Weight + Loading Quantity (unless manually overridden)
   */
  /**
   * Recalculate derived values for a loading plan row.
   * @param {HTMLTableRowElement} tr
   * @param {Object} opts
   * @param {boolean} opts.prices - also recalculate Total Ex VAT and Total Inc VAT
   */
  function recalcRow(tr, opts) {
    if (!tr) return;
    const doPrices = opts && opts.prices;

    // Get cell references
    const packagingCell = tr.querySelector(".packaging");
    const uomCell = tr.querySelector(".uom");
    const descriptionCell = tr.querySelector(".description");
    const orderQtyCell = tr.querySelector(".order-qty");
    const loadingQtyInput = tr.querySelector(".loading-qty");
    const pendingQtyCell = tr.querySelector(".pending-qty");
    const palletizedSelect = tr.querySelector(".palletized-select");
    const palletsCell = tr.querySelector(".pallets");
    const palletsWeightCell = tr.querySelector(".pallets-weight");
    const totalLitersCell = tr.querySelector(".total-liters");
    const netWeightCell = tr.querySelector(".net-weight");
    const grossWeightCell = tr.querySelector(".gross-weight");

    // Parse values
    const packaging = (packagingCell?.textContent || "").trim();
    const description = (descriptionCell?.textContent || "").trim().toUpperCase();
    const orderQty = asNum(orderQtyCell?.textContent);
    const loadingQty = asNum(loadingQtyInput?.value);
    const palletized = (palletizedSelect?.value || "No").trim().toLowerCase() === "yes";
    const numberOfPallets = asNum(palletsCell?.textContent);

    // Calculate UOM from packaging (e.g., "30x4" = 120)
    const uom = parsePackLiters(packaging);
    if (uomCell) uomCell.textContent = fmt2(uom);

    // Pending Quantity = Order Quantity − Loading Quantity
    const pendingQty = orderQty - loadingQty;
    if (pendingQtyCell) pendingQtyCell.textContent = fmt2(pendingQty);

    // Total Liters = Loading Quantity × UOM
    // Skip if user has manually overridden the value
    if (totalLitersCell && !totalLitersCell.dataset.manualOverride) {
      const totalLiters = loadingQty * uom;
      totalLitersCell.textContent = fmt2(totalLiters);
    }
    const totalLiters = asNum(totalLitersCell?.textContent);

    // Net Weight = Total Liters × DENSITY
    // DENSITY = 1.07 if description contains "COOLANT", else 0.9
    // Skip if user has manually overridden the value
    if (netWeightCell && !netWeightCell.dataset.manualOverride) {
      const density = description.includes("COOLANT") ? 1.07 : 0.9;
      const netWeight = totalLiters * density;
      netWeightCell.textContent = fmt2(netWeight);
    }
    const netWeight = asNum(netWeightCell?.textContent);

    // Pallet Weight = 0 if not palletized, else Number of Pallets × 19.38
    const palletWeight = palletized ? (numberOfPallets * 19.38) : 0;
    if (palletsWeightCell) palletsWeightCell.textContent = fmt2(palletWeight);

    // Gross Weight = Pallet Weight + Net Weight + Loading Quantity
    // Skip if user has manually overridden the value
    if (grossWeightCell && !grossWeightCell.dataset.manualOverride) {
      const grossWeight = palletWeight + netWeight + loadingQty;
      grossWeightCell.textContent = fmt2(grossWeight);
    }

    // === PRICE CALCULATIONS (only when requested) ===
    if (doPrices) {
      const valueTypeSelect = tr.querySelector(".value-type");
      const unitPriceInput = tr.querySelector(".unit-price");
      const vatCell = tr.querySelector(".vat");
      const totalExCell = tr.querySelector(".total-ex");
      const totalIncCell = tr.querySelector(".total-inc");

      const valueType = (valueTypeSelect?.value || "Sale Price").trim();
      const unitPrice = asNum(unitPriceInput?.value);
      const vat = asNum(vatCell?.textContent);

      let totalExVat = 0;
      if (valueType.toLowerCase() !== "priceless") {
        totalExVat = unitPrice * loadingQty;
      }
      if (totalExCell) totalExCell.textContent = fmt2(totalExVat);

      const totalIncVat = totalExVat + vat;
      if (totalIncCell) totalIncCell.textContent = fmt2(totalIncVat);
    }
  }

  /** Recalculate Total Inc VAT only (VAT or Total Ex changed independently) */
  function recalcTotalInc(tr) {
    if (!tr) return;
    const totalEx = asNum(tr.querySelector(".total-ex")?.textContent);
    const vat = asNum(tr.querySelector(".vat")?.textContent);
    const totalIncCell = tr.querySelector(".total-inc");
    if (totalIncCell) totalIncCell.textContent = fmt2(totalEx + vat);
  }

  // Recalculate all rows (with price recalculation)
  function recalcAllRows() {
    QA("#itemsTableBody tr.lp-data-row").forEach(tr => recalcRow(tr, { prices: true }));
  }

  // Renumber rows
  function renumberRows() {
    QA("#itemsTableBody tr.lp-data-row").forEach((tr, i) => {
      const sn = tr.querySelector(".sn");
      if (sn) sn.textContent = i + 1;
    });
  }

  // Recompute totals
  function recomputeTotals() {
    const rows = QA("#itemsTableBody tr.lp-data-row");

    let totalItems = rows.length;
    let totalOrderQty = 0;
    let totalLoadingQty = 0;
    let totalNet = 0;
    let totalGross = 0;
    let totalEx = 0;
    let totalInc = 0;
    let totalVAT = 0;

    rows.forEach(r => {
      totalOrderQty += asNum(r.querySelector(".order-qty")?.textContent);
      totalLoadingQty += asNum(r.querySelector(".loading-qty")?.value);
      totalNet += asNum(r.querySelector(".net-weight")?.textContent);
      totalGross += asNum(r.querySelector(".gross-weight")?.textContent);
      totalEx += asNum(r.querySelector(".total-ex")?.textContent);
      totalInc += asNum(r.querySelector(".total-inc")?.textContent);
      totalVAT += asNum(r.querySelector(".vat")?.textContent);
    });

    setText("#totalItems", totalItems);
    setText("#totalOrderQty", fmt2(totalOrderQty));
    setText("#totalLoadingQty", fmt2(totalLoadingQty));
    setText("#totalNetWeight", `${fmt2(totalNet)} kg`);
    setText("#totalGrossWeight", `${fmt2(totalGross)} kg`);
    setText("#totalValueExVAT", `${CURRENCY_CODE} ${fmt2(totalEx)}`);
    setText("#totalValueIncVAT", `${CURRENCY_CODE} ${fmt2(totalInc)}`);
    setText("#totalVATAmount", `${CURRENCY_CODE} ${fmt2(totalVAT)}`);

    // Backward compatibility
    setText("#totalQuantity", fmt2(totalLoadingQty));
    setText("#totalWeight", fmt2(totalGross));
    setText("#totalValue", fmt2(totalInc));

    renumberRows();
    updateAdditionalDetailsFromLoadingPlan();
  }

  // Update Additional Details from Loading Plan calculations
  function updateAdditionalDetailsFromLoadingPlan() {
    const rows = QA("#itemsTableBody tr.lp-data-row");

    let totalPallets = 0;
    let totalCartons = 0;
    let totalDrums = 0;
    let totalPails = 0;
    let totalPackages = 0;

    rows.forEach(row => {
      const packType = (row.querySelector(".pack")?.textContent || "").toUpperCase().trim();
      const loadingQty = asNum(row.querySelector(".loading-qty")?.value);

      totalPackages += loadingQty;

      if (packType.includes("CARTON") || packType.includes("CTN")) {
        totalCartons += loadingQty;
      } else if (packType.includes("DRUM")) {
        totalDrums += loadingQty;
      } else if (packType.includes("PAIL")) {
        totalPails += loadingQty;
      }

      totalPallets += asNum(row.querySelector(".pallets")?.textContent);
    });

    // Get net/gross from summary
    const summaryNetWeightEl = Q("#totalNetWeight");
    const summaryGrossWeightEl = Q("#totalGrossWeight");

    let totalNetWeight = 0;
    let totalGrossWeight = 0;

    if (summaryNetWeightEl) {
      totalNetWeight = asNum(summaryNetWeightEl.textContent.replace(/[^\d.-]/g, ''));
    }
    if (summaryGrossWeightEl) {
      totalGrossWeight = asNum(summaryGrossWeightEl.textContent.replace(/[^\d.-]/g, ''));
    }

    // Update Additional Details UI
    setText("#totalPallets", Math.round(totalPallets));
    setText("#totalCartons", Math.round(totalCartons));
    setText("#totalDrums", Math.round(totalDrums));
    setText("#totalPails", Math.round(totalPails));
    setText("#totalPackages", Math.round(totalPackages));
    setText("#totalNetWeightAD", fmt2(totalNetWeight));
    setText("#totalGrossWeightAD", fmt2(totalGrossWeight));

    // Auto-save to Dataverse if ID exists
    if (ADDITIONAL_DETAILS_ID) {
      saveAdditionalDetailsToDataverse().catch(err => {
        console.warn("Auto-save of additional details failed:", err);
      });
    }
  }

  // ── Dirty-tracking for Save All ──
  const MODIFIED_ROWS = new Set();

  function markRowDirty(tr) {
    if (!tr) return;
    const key = tr.dataset.serverId || tr.dataset.tempId || (tr.dataset.tempId = Math.random().toString(36).slice(2));
    tr.classList.add("row-modified");
    MODIFIED_ROWS.add(key);
    updateModifiedIndicator();
  }

  function updateModifiedIndicator() {
    const indicator = d.getElementById("modifiedIndicator");
    const countSpan = d.getElementById("modifiedCount");
    const saveBtn = d.getElementById("saveAllChangesBtn");
    const discardBtn = d.getElementById("discardChangesBtn");
    const count = MODIFIED_ROWS.size;
    if (indicator) indicator.style.display = count > 0 ? "inline-flex" : "none";
    if (countSpan) countSpan.textContent = count;
    if (saveBtn) saveBtn.style.display = count > 0 ? "inline-flex" : "none";
    if (discardBtn) discardBtn.style.display = count > 0 ? "inline-flex" : "none";
  }

  async function saveAllChanges() {
    if (MODIFIED_ROWS.size === 0) {
      showToast("info", "No changes to save.");
      return;
    }
    Loading.start();
    try {
      const rows = QA("#itemsTableBody tr.lp-data-row.row-modified");
      let ok = 0, fail = 0;
      for (const tr of rows) {
        try {
          await saveRowToDataverse(tr);
          tr.classList.remove("row-modified");
          ok++;
        } catch (err) {
          console.error("Failed to save row:", err);
          fail++;
        }
      }
      MODIFIED_ROWS.clear();
      updateModifiedIndicator();
      if (fail === 0) {
        showToast("success", `Saved ${ok} row${ok !== 1 ? "s" : ""} successfully.`);
      } else {
        showToast("warning", `Saved ${ok}, failed ${fail}.`);
      }
    } catch (err) {
      console.error("Save all failed:", err);
      showToast("error", "Save all failed: " + (err.message || err));
    } finally {
      Loading.stop();
    }
  }

  function discardAllChanges() {
    // Reload the page to discard all unsaved changes
    location.reload();
  }

  // Value type mapping functions
  function parseValueTypeTextToNumber(txt) {
    const v = String(txt || "").toLowerCase();
    if (v === "sale price") return 0;
    if (v === "foc value") return 1;
    if (v === "priceless") return 2;
    return 0;
  }

  function parseReleaseStatusDisplayToRaw(disp) {
    const v = String(disp || "").toUpperCase();
    if (v === "Y") return 0;
    if (v === "N") return 1;
    return 1;
  }

  // Build payload from row (matching loading_plan_view field names)
  function buildPayloadFromRow(tr, dclGuid) {
    const getNumText = sel => asNum(tr.querySelector(sel)?.textContent);
    const getNumVal = sel => asNum(tr.querySelector(sel)?.value);

    const orderNumber = (tr.querySelector(".order-no")?.textContent || "").trim();
    const itemCode = (tr.querySelector(".item-code")?.textContent || "").trim();
    const desc = (tr.querySelector(".description")?.textContent || "").trim();
    const relStatusDisp = (tr.querySelector(".release-status-select")?.value || "N").trim();
    const packaging = (tr.querySelector(".packaging")?.textContent || "").trim();
    const uomNumeric = (tr.querySelector(".uom")?.textContent || "").trim();
    const packType = (tr.querySelector(".pack")?.textContent || "").trim();

    const ordQty = getNumText(".order-qty");
    const loadQty = getNumVal(".loading-qty");
    const pendQty = getNumText(".pending-qty");

    const palletizedSel = (tr.querySelector(".palletized-select")?.value || "No").trim();
    const palletCount = getNumText(".pallets");
    const palletWeight = getNumText(".pallets-weight");

    const totalVol = getNumText(".total-liters");
    const netW = getNumText(".net-weight");
    const grossW = getNumText(".gross-weight");

    const valueTypeLbl = (tr.querySelector(".value-type")?.value || "Sale Price").trim();
    const unitPrice = getNumVal(".unit-price");
    const vat = getNumText(".vat");
    const totalEx = getNumText(".total-ex");
    const totalInc = getNumText(".total-inc");

    const containerNo = (tr.querySelector(".container-no")?.value || "").trim();
    const hsCode = (tr.querySelector(".hs-code")?.value || "").trim();

    const payload = {
      cr650_ordernumber: orderNumber,
      cr650_itemcode: itemCode,
      cr650_itemdescription: desc,
      cr650_releasestatus: parseReleaseStatusDisplayToRaw(relStatusDisp),
      cr650_packagingdetails: packaging,
      cr650_unitofmeasure: uomNumeric,
      cr650_packagetype: packType,
      cr650_orderedquantity: ordQty,
      cr650_loadedquantity: loadQty,
      cr650_pendingquantity: pendQty,
      cr650_ispalletized: palletizedSel.toLowerCase() === "yes",
      cr650_palletcount: palletCount,
      cr650_palletweight: palletWeight,
      cr650_totalvolumeorweight: totalVol,
      cr650_netweightkg: netW,
      cr650_grossweightkg: grossW,
      cr650_valuetype: parseValueTypeTextToNumber(valueTypeLbl),
      cr650_unitprice: unitPrice,
      cr650_vat5percent: vat,
      cr650_totalexcludingvat: totalEx,
      cr650_totalincludingvat: totalInc,
      cr650_containernumber: containerNo,
      cr650_hscode: hsCode
    };

    if (dclGuid && guidOK(dclGuid)) {
      payload["cr650_dcl_number@odata.bind"] = `/cr650_dcl_masters(${dclGuid})`;
    }

    return payload;
  }

  // Save row to Dataverse
  async function saveRowToDataverse(tr) {
    const serverId = tr.dataset.serverId;
    if (!serverId) {
      // No server ID - create new record
      return await createServerRowFromTr(tr, CURRENT_DCL_ID);
    }

    const payload = buildPayloadFromRow(tr, null); // Don't include DCL bind for PATCH

    try {
      await safeAjaxWithToken({
        type: "PATCH",
        url: `${API_URL_LOADPLANS}(${serverId})`,
        data: JSON.stringify(payload),
        contentType: "application/json; charset=utf-8",
        headers: { "If-Match": "*" }
      });
      return true;
    } catch (err) {
      console.error("Failed to save row:", err);
      return false;
    }
  }

  // Create new row in Dataverse
  async function createServerRowFromTr(tr, dclGuid) {
    const payload = buildPayloadFromRow(tr, dclGuid);

    try {
      const res = await safeAjaxWithToken({
        type: "POST",
        url: API_URL_LOADPLANS,
        data: JSON.stringify(payload),
        contentType: "application/json; charset=utf-8",
        headers: {
          Accept: "application/json;odata.metadata=minimal",
          Prefer: "return=representation"
        }
      });

      let newId = res && (res.cr650_dcl_loading_planid || res.id);

      // If ID not in response, fetch it back
      if (!newId) {
        await new Promise(resolve => setTimeout(resolve, 800));

        const orderNo = (tr.querySelector(".order-no")?.textContent || "").trim();
        const itemCode = (tr.querySelector(".item-code")?.textContent || "").trim();

        if (orderNo && itemCode && dclGuid) {
          const escapedOrder = orderNo.replace(/'/g, "''");
          const escapedItem = itemCode.replace(/'/g, "''");

          const fetchUrl = `${API_URL_LOADPLANS}?$filter=_cr650_dcl_number_value eq ${dclGuid} and cr650_ordernumber eq '${escapedOrder}' and cr650_itemcode eq '${escapedItem}'&$orderby=createdon desc&$top=1&$select=cr650_dcl_loading_planid`;

          const fetchRes = await $.ajax({
            type: "GET",
            url: fetchUrl,
            headers: { Accept: "application/json;odata.metadata=minimal" },
            dataType: "json"
          });

          if (fetchRes && fetchRes.value && fetchRes.value.length > 0) {
            newId = fetchRes.value[0].cr650_dcl_loading_planid;
          }
        }
      }

      if (newId) {
        tr.dataset.serverId = newId;
      }
      return true;
    } catch (err) {
      console.error("Failed to create row:", err);
      return false;
    }
  }

  // Load order items
  async function loadOrderItems() {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) return [];
    CURRENT_DCL_ID = dclGuid;

    try {
      const url = `${API_URL_LOADPLANS}?$filter=_cr650_dcl_number_value eq ${dclGuid}&$orderby=cr650_ordernumber`;
      const results = await $.ajax({
        type: "GET",
        url: url,
        headers: { Accept: "application/json;odata.metadata=minimal" },
        dataType: "json"
      });
      return results.value || [];
    } catch (err) {
      console.error("Error loading order items:", err);
      return [];
    }
  }

  // Render order items
  function renderOrderItems(items) {
    const tbody = Q("#itemsTableBody");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (!items || items.length === 0) {
      tbody.innerHTML = `
        <tr class="empty-row">
          <td colspan="24">
            <i class="fas fa-inbox"></i>
            No order items found
          </td>
        </tr>
      `;
      recomputeTotals();
      return;
    }

    items.forEach((raw, i) => {
      const item = parseLpRow(raw);
      const tr = makeRowEl(item, i);
      tbody.appendChild(tr);
    });

    // Don't call recalcAllRows() here - we're loading saved values from Dataverse
    // and want to preserve user's manually entered values (like gross weight)
    recomputeTotals();
    attachRowEventHandlers();
  }

  // Attach event handlers to rows
  function attachRowEventHandlers() {
    // ── Value Type change → recalc Total Ex + Total Inc ──
    d.addEventListener("change", e => {
      if (e.target.matches(".value-type")) {
        const tr = e.target.closest("tr.lp-data-row");
        if (tr) {
          recalcRow(tr, { prices: true });
          recomputeTotals();
          markRowDirty(tr);
        }
      }
    });

    // ── Unit Price input → recalc Total Ex + Total Inc ──
    d.addEventListener("input", e => {
      if (e.target.matches(".unit-price")) {
        const tr = e.target.closest("tr.lp-data-row");
        if (tr) {
          recalcRow(tr, { prices: true });
          recomputeTotals();
          markRowDirty(tr);
        }
      }
    });

    // ── VAT cell edited → only recalc Total Inc ──
    d.addEventListener("input", e => {
      const td = e.target.closest("td.vat");
      if (!td || !td.isContentEditable) return;
      const tr = td.closest("tr.lp-data-row");
      if (tr) {
        recalcTotalInc(tr);
        recomputeTotals();
        markRowDirty(tr);
      }
    });

    // ── Total Ex edited → only recalc Total Inc ──
    d.addEventListener("input", e => {
      const td = e.target.closest("td.total-ex");
      if (!td || !td.isContentEditable) return;
      const tr = td.closest("tr.lp-data-row");
      if (tr) {
        recalcTotalInc(tr);
        recomputeTotals();
        markRowDirty(tr);
      }
    });

    // ── Total Inc edited directly → just mark dirty, no recalc ──
    d.addEventListener("input", e => {
      const td = e.target.closest("td.total-inc");
      if (!td || !td.isContentEditable) return;
      const tr = td.closest("tr.lp-data-row");
      if (tr) {
        recomputeTotals();
        markRowDirty(tr);
      }
    });

    // ── Container # or HS Code change → mark dirty ──
    d.addEventListener("change", e => {
      if (e.target.matches(".container-no, .hs-code")) {
        const tr = e.target.closest("tr.lp-data-row");
        if (tr) markRowDirty(tr);
      }
    });

    // ── Save All button ──
    const saveBtn = d.getElementById("saveAllChangesBtn");
    if (saveBtn) saveBtn.addEventListener("click", saveAllChanges);

    // ── Discard button ──
    const discardBtn = d.getElementById("discardChangesBtn");
    if (discardBtn) discardBtn.addEventListener("click", () => {
      if (confirm("Discard all unsaved changes?")) discardAllChanges();
    });

    // ── Ctrl+S shortcut ──
    d.addEventListener("keydown", e => {
      if ((e.ctrlKey || e.metaKey) && e.key === "s") {
        e.preventDefault();
        saveAllChanges();
      }
    });
  }

  // Initialize order items
  async function initOrderItems() {
    try {
      Loading.start();

      // Currency code is already set by fetchDclMaster in main init

      const items = await loadOrderItems();
      renderOrderItems(items);

      // Fetch shipped data for container # dropdown
      if (items && items.length > 0) {
        await fetchShippedForAllOrders(items);
        populateAllContainerDropdowns();
      }

      // Fetch and populate HS codes
      await fetchAndCacheHSCodes();
      populateAllHSCodes();
    } catch (err) {
      console.error("initOrderItems failed:", err);
      const tbody = Q("#itemsTableBody");
      if (tbody) {
        tbody.innerHTML = `
          <tr class="empty-row">
            <td colspan="24">
              <i class="fas fa-exclamation-triangle"></i>
              Error loading order items
            </td>
          </tr>
        `;
      }
    } finally {
      Loading.stop();
    }

    // Update All button handler - Fetch AR Report and update rows
    const updateAllBtn = Q("#updateAllBtn");
    if (updateAllBtn) {
      updateAllBtn.addEventListener("click", async () => {
        Loading.start();
        try {
          // STEP 1: Validate DCL ID
          const dclId = getQueryId();
          if (!dclId || !guidOK(dclId)) {
            showToast("error", "No valid DCL ID found");
            return;
          }

          console.log("📋 Update All: Starting for DCL", dclId);

          // STEP 2: Get Customer Number from DCL Master
          const customerNumber = await getCustomerNumberFromDcl(dclId);
          if (!customerNumber) {
            showToast("error", "Could not find customer number for this DCL");
            return;
          }
          console.log("📋 Customer Number:", customerNumber);

          // STEP 3: Fetch AR Report by Customer
          const arReportData = await fetchARReportByCustomer(customerNumber);
          console.log(`📋 Fetched ${arReportData.length} AR Report records`);

          if (!arReportData.length) {
            showToast("warning", "No AR Report data found for this customer");
            // Still save any existing changes
            recalcAllRows();
            recomputeTotals();
            return;
          }

          // STEP 4: Build AR Index (lookup map)
          const arIndex = buildARIndexByOrderAndItem(arReportData);
          console.log(`📋 Built AR Index with ${arIndex.size} unique order/item combinations`);

          // STEP 5: Loop through each Loading Plan row
          const rows = QA("#itemsTableBody tr.lp-data-row");
          let updatedCount = 0;
          let notFoundCount = 0;
          let errorCount = 0;

          for (const tr of rows) {
            // Extract order number and item code from the row
            const orderNo = (tr.querySelector(".order-no")?.textContent || "").trim();
            const itemCode = (tr.querySelector(".item-code")?.textContent || "").trim();

            if (!orderNo || !itemCode) continue;

            // Build lookup key
            const key = `${orderNo}|${itemCode}`;

            // Look up in AR Index
            const arData = arIndex.get(key);

            if (arData) {
              // FOUND - Update the row
              const success = await updateRowFromARData(tr, arData);

              if (success) {
                updatedCount++;
                // Flash green for visual feedback
                tr.style.background = "#d4edda";
                setTimeout(() => { tr.style.background = ""; }, 1500);
              } else {
                errorCount++;
              }
            } else {
              // NOT FOUND in AR Report
              notFoundCount++;
              console.log(`⚠️ No AR data for: ${key}`);
            }
          }

          // STEP 6: Recalculate All
          recalcAllRows();
          recomputeTotals();

          // STEP 7: Show Summary Alert
          let message = `Updated: ${updatedCount} items`;
          if (notFoundCount > 0) {
            message += `, Not found in AR: ${notFoundCount}`;
          }
          if (errorCount > 0) {
            message += `, Errors: ${errorCount}`;
          }

          if (errorCount > 0) {
            showToast("warning", message);
          } else if (updatedCount > 0) {
            showToast("success", message);
          } else {
            showToast("info", message);
          }

          console.log("📋 Update All completed:", { updatedCount, notFoundCount, errorCount });

        } catch (err) {
          console.error("Update All failed:", err);
          showToast("error", "Failed to update all items: " + (err.message || err));
        } finally {
          Loading.stop();
        }
      });
    }

    // Fullscreen button functionality
    const fullscreenBtn = Q("#fullscreenBtn");
    const fullscreenOverlay = Q("#fullscreenOverlay");
    const closeFullscreenBtn = Q("#closeFullscreenBtn");
    const fsCloseBtn = Q("#fsCloseBtn");
    const fsUpdateAllBtn = Q("#fsUpdateAllBtn");
    const itemsTableContainer = Q("#itemsTableContainer");
    const fullscreenTableContainer = Q("#fullscreenTableContainer");

    let originalTableParent = null;

    function openFullscreen() {
      if (!fullscreenOverlay || !itemsTableContainer || !fullscreenTableContainer) return;

      // Store original parent
      originalTableParent = itemsTableContainer.parentNode;

      // Move table to fullscreen container
      fullscreenTableContainer.appendChild(itemsTableContainer);

      // Show overlay
      fullscreenOverlay.classList.add("active");
      d.body.classList.add("fullscreen-active");

      // Focus on overlay for accessibility
      fullscreenOverlay.focus();
    }

    function closeFullscreen() {
      if (!fullscreenOverlay || !itemsTableContainer || !originalTableParent) return;

      // Move table back to original location
      const itemsSection = Q(".items-section");
      const itemsHeader = Q(".items-header");
      if (itemsHeader && itemsHeader.nextSibling) {
        itemsSection.insertBefore(itemsTableContainer, itemsHeader.nextSibling);
      } else if (itemsSection) {
        itemsSection.appendChild(itemsTableContainer);
      }

      // Hide overlay
      fullscreenOverlay.classList.remove("active");
      d.body.classList.remove("fullscreen-active");
    }

    // Fullscreen button click
    if (fullscreenBtn) {
      fullscreenBtn.addEventListener("click", openFullscreen);
    }

    // Close buttons
    if (closeFullscreenBtn) {
      closeFullscreenBtn.addEventListener("click", closeFullscreen);
    }
    if (fsCloseBtn) {
      fsCloseBtn.addEventListener("click", closeFullscreen);
    }

    // ESC key to close fullscreen
    d.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && fullscreenOverlay?.classList.contains("active")) {
        closeFullscreen();
      }
    });

    // Click outside to close (on overlay background)
    if (fullscreenOverlay) {
      fullscreenOverlay.addEventListener("click", (e) => {
        if (e.target === fullscreenOverlay) {
          closeFullscreen();
        }
      });
    }

    // Fullscreen Update All button
    if (fsUpdateAllBtn) {
      fsUpdateAllBtn.addEventListener("click", () => {
        if (updateAllBtn) updateAllBtn.click();
      });
    }
  }

  /* =============================
     ADDITIONAL DETAILS MODULE (Full Features - Same as Loading Plan View)
     ============================= */

  const API_URL_ADDITIONAL = "/_api/cr650_dcl_additional_detailses";

  // Get or create Additional Details record
  async function getOrCreateAdditionalDetailsRecord(dclId) {
    const url = `${API_URL_ADDITIONAL}?$select=cr650_dcl_additional_detailsid,cr650_additionalcomments,cr650_printon,cr650_hidden_details&$filter=_cr650_dclreference_value eq ${dclId}`;

    const res = await $.ajax({
      type: "GET",
      url: url,
      headers: { Accept: "application/json;odata.metadata=minimal" },
      dataType: "json"
    });

    if (res.value?.length > 0) {
      const record = res.value[0];
      ADDITIONAL_DETAILS_ID = record.cr650_dcl_additional_detailsid || null;
      return record;
    }

    // Create new record
    const createBody = {
      "cr650_name": "Additional Details",
      "cr650_additionalcomments": "",
      "cr650_printon": null,
      "cr650_hidden_details": "",
      "cr650_dclreference@odata.bind": `/cr650_dcl_masters(${dclId})`
    };

    const created = await safeAjaxWithToken({
      type: "POST",
      url: API_URL_ADDITIONAL,
      data: JSON.stringify(createBody),
      contentType: "application/json; charset=utf-8",
      headers: { Accept: "application/json;odata.metadata=minimal", "Prefer": "return=representation" }
    });

    if (created && created.cr650_dcl_additional_detailsid) {
      ADDITIONAL_DETAILS_ID = created.cr650_dcl_additional_detailsid;
    }

    return created;
  }

  // Save Additional Details to Dataverse
  async function saveAdditionalDetailsToDataverse() {
    if (!CURRENT_DCL_ID || !guidOK(CURRENT_DCL_ID) || !ADDITIONAL_DETAILS_ID) {
      return;
    }

    try {
      // Hidden details
      const hiddenStr = Array.from(hiddenDetails).join(",");

      // Print options
      const clCheckbox = Q("#printAdditionalOnCL");
      const plCheckbox = Q("#printAdditionalOnPL");
      const bothCheckbox = Q("#printAdditionalOnBoth");

      let printOn = null;
      if (bothCheckbox && bothCheckbox.checked) {
        printOn = 3;
      } else if (clCheckbox && clCheckbox.checked) {
        printOn = 1;
      } else if (plCheckbox && plCheckbox.checked) {
        printOn = 2;
      }

      // Get calculated totals from UI
      const getText = id => {
        const el = Q(`#${id}`);
        if (!el) return 0;
        return asNum(el.textContent.replace(/[^\d.-]/g, ""));
      };

      const payload = {
        cr650_hidden_details: hiddenStr || null,
        cr650_printon: printOn,
        cr650_totalpallets: getText("totalPallets"),
        cr650_totalcartons: getText("totalCartons"),
        cr650_totaldrums: getText("totalDrums"),
        cr650_totalpails: getText("totalPails"),
        cr650_totalpackages: getText("totalPackages"),
        cr650_totalnetweight: getText("totalNetWeightAD"),
        cr650_totalgrossweight: getText("totalGrossWeightAD")
      };

      await safeAjaxWithToken({
        type: "PATCH",
        url: `${API_URL_ADDITIONAL}(${ADDITIONAL_DETAILS_ID})`,
        data: JSON.stringify(payload),
        contentType: "application/json; charset=utf-8",
        headers: { "If-Match": "*" }
      });
    } catch (err) {
      console.error("Failed to save additional details:", err);
    }
  }

  // Update Add Detail dropdown
  function updateAddDetailDropdown() {
    const select = Q("#addDetailSelect");
    const addBtn = Q("#addDetailBtn");
    if (!select) return;

    select.innerHTML = '<option value="">Select detail to add back</option>';

    Object.entries(DETAIL_LABELS).forEach(([key, label]) => {
      if (hiddenDetails.has(key)) {
        const opt = d.createElement("option");
        opt.value = key;
        opt.textContent = label;
        select.appendChild(opt);
      }
    });

    if (addBtn) {
      addBtn.disabled = select.options.length <= 1;
    }
  }

  // Remove detail row
  function removeDetailRow(btn, key) {
    const tr = btn.closest("tr");
    if (!tr) return;

    const keyAttr = tr.getAttribute("data-detail") || key || "";
    if (!keyAttr) return;

    hiddenDetails.add(keyAttr);
    tr.style.display = "none";

    updateAddDetailDropdown();
    saveAdditionalDetailsToDataverse().catch(err => {
      console.error("Failed to save after remove:", err);
    });
  }
  w.removeDetailRow = removeDetailRow;

  // Add removed detail back
  function addRemovedDetailBack() {
    const select = Q("#addDetailSelect");
    if (!select) return;

    const key = select.value;
    if (!key) return;

    const row = d.querySelector(`tr[data-detail="${key}"]`);
    if (row) {
      row.style.display = "";
    }

    hiddenDetails.delete(key);
    select.value = "";

    updateAddDetailDropdown();
    saveAdditionalDetailsToDataverse().catch(err => {
      console.error("Failed to save after add back:", err);
    });
  }

  // Load Additional Details
  async function loadAdditionalDetails() {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) return null;

    try {
      const record = await getOrCreateAdditionalDetailsRecord(dclGuid);

      if (record) {
        // Handle hidden details
        const hiddenStr = record.cr650_hidden_details || "";
        hiddenStr.split(",").map(s => s.trim()).filter(Boolean).forEach(key => {
          hiddenDetails.add(key);
          const row = d.querySelector(`tr[data-detail="${key}"]`);
          if (row) row.style.display = "none";
        });

        // Handle print options
        const printOn = record.cr650_printon;
        const clCheckbox = Q("#printAdditionalOnCL");
        const plCheckbox = Q("#printAdditionalOnPL");
        const bothCheckbox = Q("#printAdditionalOnBoth");

        if (printOn === 3 && bothCheckbox) {
          bothCheckbox.checked = true;
        } else if (printOn === 1 && clCheckbox) {
          clCheckbox.checked = true;
        } else if (printOn === 2 && plCheckbox) {
          plCheckbox.checked = true;
        }

        updateAddDetailDropdown();
      }

      return record;
    } catch (err) {
      console.error("Error loading additional details:", err);
      return null;
    }
  }

  // Initialize Additional Details
  async function initAdditionalDetails() {
    try {
      await loadAdditionalDetails();

      // Add Detail button handler
      const addDetailBtn = Q("#addDetailBtn");
      if (addDetailBtn) {
        addDetailBtn.addEventListener("click", addRemovedDetailBack);
      }

      // Print options change handlers
      const printCheckboxes = QA("#printAdditionalOnCL, #printAdditionalOnPL, #printAdditionalOnBoth");
      printCheckboxes.forEach(cb => {
        cb.addEventListener("change", () => {
          // Uncheck others when "Both" is checked
          if (cb.id === "printAdditionalOnBoth" && cb.checked) {
            Q("#printAdditionalOnCL").checked = false;
            Q("#printAdditionalOnPL").checked = false;
          } else if (cb.checked) {
            Q("#printAdditionalOnBoth").checked = false;
          }

          saveAdditionalDetailsToDataverse().catch(err => {
            console.error("Failed to save print options:", err);
          });
        });
      });

    } catch (err) {
      console.error("initAdditionalDetails failed:", err);
    }
  }

  // Expose functions globally
  w.initOrderItems = initOrderItems;
  w.loadOrderItems = loadOrderItems;
  w.renderOrderItems = renderOrderItems;
  w.recomputeTotals = recomputeTotals;
  w.recalcAllRows = recalcAllRows;
  w.initAdditionalDetails = initAdditionalDetails;
  w.loadAdditionalDetails = loadAdditionalDetails;
  w.removeDetailRow = removeDetailRow;
  w.addRemovedDetailBack = addRemovedDetailBack;
  w.saveAdditionalDetailsToDataverse = saveAdditionalDetailsToDataverse;

  // Helper function for toast notifications
  function showToast(type, message) {
    const toast = d.createElement("div");
    toast.className = `toast ${type}`;
    toast.textContent = message;
    d.body.appendChild(toast);
    setTimeout(() => toast.classList.add("show"), 10);
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }

  // Expose functions globally
  w.initDiscountCharges = initDiscountCharges;
  w.saveAllDiscountsAndCharges = saveAllDiscountsAndCharges;
  w.loadDiscountsFromDataverse = loadDiscountsFromDataverse;
  w.renderDiscountCharges = renderDiscountCharges;
  w.updateDiscountChargeSummary = updateDiscountChargeSummary;

  /* =============================
     TERMS & CONDITIONS MODULE
     ============================= */

  const TERMS_FIELDS = [
    "cr650_dcl_terms_conditionsid",
    "cr650_termcondition",
    "cr650_is_printed_ci",
    "cr650_is_printed_pl",
    "createdon",
    "_cr650_dcl_number_value"
  ];

  function buildTermsQueryUrl(dclGuid) {
    const select = "$select=" + encodeURIComponent(TERMS_FIELDS.join(","));
    const filter = "$filter=" + encodeURIComponent(`_cr650_dcl_number_value eq ${dclGuid}`);
    const orderby = "$orderby=" + encodeURIComponent("createdon asc");
    return `${API_URL_TERMS}?${select}&${filter}&${orderby}`;
  }

  async function loadTermsConditions(dclGuid) {
    if (!dclGuid) return [];
    const records = [];

    async function loop(nextUrl) {
      const data = await safeAjaxWithToken({ type: "GET", url: nextUrl });
      const rows = (data && (data.value || data)) || [];
      rows.forEach(r => records.push(r));
      if (data && data["@odata.nextLink"]) {
        return loop(data["@odata.nextLink"]);
      }
    }

    await loop(buildTermsQueryUrl(dclGuid));
    return records.map(normalizeTermsCondition);
  }

  function normalizeTermsCondition(r) {
    return {
      id: r.cr650_dcl_terms_conditionsid || "",
      text: r.cr650_termcondition || "",
      printOnCI: !!r.cr650_is_printed_ci,
      printOnPL: !!r.cr650_is_printed_pl,
      created: r.createdon || "",
      dclGuid: r._cr650_dcl_number_value || ""
    };
  }

  async function createTermsCondition(textVal, printCI, printPL) {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) throw new Error("No DCL GUID available");

    const bodyObj = {
      cr650_termcondition: textVal,
      cr650_is_printed_ci: !!printCI,
      cr650_is_printed_pl: !!printPL,
      "cr650_dcl_number@odata.bind": "/cr650_dcl_masters(" + dclGuid + ")"
    };

    await safeAjaxWithToken({
      type: "POST",
      url: API_URL_TERMS,
      headers: {
        "Content-Type": "application/json; charset=utf-8"
      },
      data: JSON.stringify(bodyObj)
    });

    // Return success - we'll refresh the list to get the created record
    return { ok: true };
  }

  async function updateTermsCondition(id, textVal, printCI, printPL) {
    const url = `${API_URL_TERMS}(${encodeURIComponent(id)})`;
    const bodyObj = {
      cr650_termcondition: textVal,
      cr650_is_printed_ci: !!printCI,
      cr650_is_printed_pl: !!printPL
    };

    await safeAjaxWithToken({
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

  async function deleteTermsCondition(id) {
    const url = `${API_URL_TERMS}(${encodeURIComponent(id)})`;

    await safeAjaxWithToken({
      type: "DELETE",
      url,
      headers: { "If-Match": "*" }
    });

    return { ok: true };
  }

  function escapeHtmlTC(str) {
    if (!str) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function buildTermsConditionItem(item, isNew, index) {
    const wrapper = d.createElement("div");
    wrapper.className = "terms-condition-item";
    wrapper.dataset.termsId = item.id || "";

    const createdDate = item.created ? new Date(item.created) : null;
    const humanTime = (createdDate && !isNaN(createdDate))
      ? createdDate.toLocaleString()
      : "";

    wrapper.innerHTML = `
      <div class="form-group">
        <label for="termsComments${index}">Terms & Conditions</label>
        <textarea
          id="termsComments${index}"
          class="terms-textarea"
          rows="4"
          placeholder="Enter terms and conditions..."
        >${escapeHtmlTC(item.text || "")}</textarea>
      </div>

      <div class="form-group">
        <span style="font-weight: 600; color: #374151;">Print on:</span>
        <div class="print-options" style="display:flex; gap:1.5rem; margin-top:0.5rem;">
          <label style="display:flex; align-items:center; gap:0.5rem; cursor:pointer;">
            <input type="checkbox" class="tc-print-ci" ${item.printOnCI ? "checked" : ""}>
            <span>Commercial Invoice (CI)</span>
          </label>
          <label style="display:flex; align-items:center; gap:0.5rem; cursor:pointer;">
            <input type="checkbox" class="tc-print-pl" ${item.printOnPL ? "checked" : ""}>
            <span>Packing List (PL)</span>
          </label>
        </div>
      </div>

      <div class="notify-actions">
        <span class="notify-meta">
          ${humanTime ? `Saved: ${humanTime}` : "New entry"}
        </span>
        <div class="notify-buttons">
          <button
            type="button"
            class="tc-save-btn"
            data-mode="${isNew ? "create" : "update"}"
          >Save</button>
          <button
            type="button"
            class="tc-delete-btn ${isNew ? "hidden" : ""}"
          >Delete</button>
        </div>
      </div>
    `;

    return wrapper;
  }

  function renderTermsConditionsList(list) {
    const container = d.getElementById("termsConditionsContainer");
    if (!container) return;
    container.innerHTML = "";

    if (!list.length) {
      const draftItem = {
        id: "",
        text: "",
        printOnCI: false,
        printOnPL: false,
        created: "",
        dclGuid: getQueryId() || ""
      };
      const rowEl = buildTermsConditionItem(draftItem, true, 1);
      container.appendChild(rowEl);
      return;
    }

    list.forEach((item, idx) => {
      const rowEl = buildTermsConditionItem(item, false, idx + 1);
      container.appendChild(rowEl);
    });
  }

  function addTermsCondition() {
    const container = d.getElementById("termsConditionsContainer");
    if (!container) return;

    // Don't add if there's already an unsaved draft
    if (container.querySelector('.terms-condition-item[data-terms-id=""]')) {
      showToast("warning", "Please save the current draft first");
      return;
    }

    const draftItem = {
      id: "",
      text: "",
      printOnCI: false,
      printOnPL: false,
      created: "",
      dclGuid: getQueryId() || ""
    };
    const idx = container.children.length + 1;
    const rowEl = buildTermsConditionItem(draftItem, true, idx);
    container.appendChild(rowEl);
  }

  async function initTermsConditions() {
    const dclGuid = getQueryId();
    if (!dclGuid || !guidOK(dclGuid)) {
      console.warn("No valid DCL GUID for Terms & Conditions");
      const container = Q("#termsConditionsContainer");
      if (container) {
        container.innerHTML = `<div class="terms-loading" style="color: var(--neutral-gray);"><i class="fas fa-info-circle"></i> Select a DCL to view terms & conditions</div>`;
      }
      return;
    }

    try {
      // Load existing terms
      const termsList = await loadTermsConditions(dclGuid);
      renderTermsConditionsList(termsList);

      // Add button handler
      const addBtn = Q("#addTermsConditionBtn");
      if (addBtn) {
        addBtn.addEventListener("click", () => {
          addTermsCondition();
        });
      }

      // Event delegation for save/delete buttons
      const container = Q("#termsConditionsContainer");
      if (container) {
        container.addEventListener("click", async (e) => {
          // Save button
          if (e.target.classList.contains("tc-save-btn")) {
            const card = e.target.closest(".terms-condition-item");
            if (!card) return;

            const textarea = card.querySelector(".terms-textarea");
            const cbCI = card.querySelector(".tc-print-ci");
            const cbPL = card.querySelector(".tc-print-pl");

            const textVal = (textarea?.value || "").trim();
            const printCI = !!cbCI?.checked;
            const printPL = !!cbPL?.checked;

            const id = card.dataset.termsId || "";
            const mode = e.target.getAttribute("data-mode") || (id ? "update" : "create");

            if (!textVal) {
              if (!id) card.remove();
              showToast("warning", "Please enter terms text");
              return;
            }

            Loading.start();
            try {
              if (mode === "create" || !id) {
                await createTermsCondition(textVal, printCI, printPL);
                showToast("success", "Terms & Conditions created");
              } else {
                await updateTermsCondition(id, textVal, printCI, printPL);
                showToast("success", "Terms & Conditions updated");
              }

              const refreshedTerms = await loadTermsConditions(dclGuid);
              renderTermsConditionsList(refreshedTerms);
            } catch (err) {
              console.error("Failed to save terms & conditions", err);
              showToast("error", "Failed to save Terms & Conditions");
            } finally {
              Loading.stop();
            }
            return;
          }

          // Delete button
          if (e.target.classList.contains("tc-delete-btn")) {
            const card = e.target.closest(".terms-condition-item");
            if (!card) return;

            const id = card.dataset.termsId || "";

            if (!id) {
              card.remove();
              return;
            }

            const sure = confirm("Delete this Terms & Conditions entry?");
            if (!sure) return;

            Loading.start();
            try {
              await deleteTermsCondition(id);
              const refreshedTerms = await loadTermsConditions(dclGuid);
              renderTermsConditionsList(refreshedTerms);
              showToast("success", "Terms & Conditions deleted");
            } catch (err) {
              console.error("Failed to delete terms & conditions", err);
              showToast("error", "Failed to delete Terms & Conditions");
            } finally {
              Loading.stop();
            }
            return;
          }
        });
      }

    } catch (err) {
      console.error("initTermsConditions failed:", err);
      const container = Q("#termsConditionsContainer");
      if (container) {
        container.innerHTML = `<div class="terms-loading" style="color: var(--error-red);"><i class="fas fa-exclamation-circle"></i> Failed to load terms & conditions</div>`;
      }
    }
  }

  // Expose Terms & Conditions functions globally
  w.initTermsConditions = initTermsConditions;
  w.loadTermsConditions = loadTermsConditions;
  w.renderTermsConditionsList = renderTermsConditionsList;
  w.addTermsCondition = addTermsCondition;

  /* =========================
     4) DATAVERSE FETCH (shared)
     ========================= */
  async function fetchMaster(dclId) {
    const sel = Object.values(MF).join(",");  // ✅ This now includes status
    for (const f of [`${MF.id} eq guid('${dclId}')`, `${MF.id} eq '${dclId}'`]) {
      const url = `${API_URL_MASTER}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$top=1`;
      try {
        const r = await safeAjax(url);
        if (r?.value?.length) {
          const record = r.value[0];

          // ✅ NEW: Capture status
          DCL_STATUS = record[MF.status] || null;
          console.log("📋 DCL Status captured:", DCL_STATUS);

          return record;
        }
      } catch { }
    }
    throw new Error("DCL not found");
  }
  async function fetchContainers(dclId) {
    const sel = [CF.id, CF.dclLkp, CF.code, CF.type, CF.weight, CF.totalPkgs, CF.totalGross, CF.totalNet].join(",");
    for (const f of [`${CF.dclLkp} eq guid('${dclId}')`, `${CF.dclLkp} eq '${dclId}'`]) {
      const url = `${API_URL_CONTAINERS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}`;
      try { const r = await safeAjax(url); if (Array.isArray(r?.value)) return r.value; } catch { }
    }
    return [];
  }
  async function fetchItemsByContainer(cid) {
    const sel = [IF.id, IF.containerLkp, IF.planLkp, IF.qty].join(",");
    for (const f of [`${IF.containerLkp} eq guid('${cid}')`, `${IF.containerLkp} eq '${cid}'`]) {
      const url = `${API_URL_ITEMS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}`;
      try { const r = await safeAjax(url); if (Array.isArray(r?.value)) return r.value; } catch { }
    }
    return [];
  }
  async function fetchPlanCI(pid) {
    const L = LF_CI;
    const sel = [L.id, L.orderNo, L.itemCode, L.hsCode, L.description, L.packaging, L.unitPrice, L.vat5, L.totalInclVat].join(",");
    for (const f of [`${L.id} eq guid('${pid}')`, `${L.id} eq '${pid}'`]) {
      const url = `${API_URL_LOADPLANS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$top=1`;
      try { const r = await safeAjax(url); if (r?.value?.length) return r.value[0]; } catch { }
    }
    return null;
  }
  async function fetchPlanPL(pid) {
    const L = LF_PL;
    const sel = [L.id, L.orderNo, L.itemCode, L.hsCode, L.description, L.packaging, L.uom, L.totalLiters, L.netKg, L.grossKg, L.loadedQty, L.containerNumber, L.palletWeight].join(",");  // ✅ Added L.hsCode
    for (const f of [`${L.id} eq guid('${pid}')`, `${L.id} eq '${pid}'`]) {
      const url = `${API_URL_LOADPLANS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$top=1`;
      try { const r = await safeAjax(url); if (r?.value?.length) return r.value[0]; } catch { }
    }
    return null;
  }
  async function fetchLoadingPlansForCI(dclId) {
    const L = LF_CI;
    const sel = [L.id, L.orderNo, L.itemCode, L.hsCode, L.description, L.packaging, L.unitPrice, L.vat5, L.totalInclVat, "cr650_loadedquantity"].join(",");
    for (const f of [`_cr650_dcl_number_value eq guid('${dclId}')`, `_cr650_dcl_number_value eq '${dclId}'`]) {
      const url = `${API_URL_LOADPLANS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$orderby=cr650_ordernumber`;
      try { const r = await safeAjax(url); if (Array.isArray(r?.value)) return r.value; } catch { }
    }
    return [];
  }
  async function fetchLoadingPlansForPL(dclId) {
    const L = LF_PL;
    const sel = [L.id, L.orderNo, L.itemCode, L.hsCode, L.description, L.packaging, L.uom, L.totalLiters, L.netKg, L.grossKg, L.loadedQty, L.containerNumber, L.palletWeight].join(",");
    for (const f of [`_cr650_dcl_number_value eq guid('${dclId}')`, `_cr650_dcl_number_value eq '${dclId}'`]) {
      const url = `${API_URL_LOADPLANS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$orderby=cr650_ordernumber`;
      try { const r = await safeAjax(url); if (Array.isArray(r?.value)) return r.value; } catch { }
    }
    return [];
  }
  async function fetchBrands(dclId) {
    const sel = ["cr650_dcl_brandid", "cr650_brand", "_cr650_dcl_number_value"].join(",");
    for (const f of [`_cr650_dcl_number_value eq guid('${dclId}')`, `_cr650_dcl_number_value eq '${dclId}'`]) {
      const url = `${API_URL_BRANDS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}`;
      try {
        const r = await safeAjax(url);
        if (Array.isArray(r?.value) && r.value.length) {
          const brands = r.value.map(x => x.cr650_brand).filter(Boolean);
          if (brands.length) return brands;  // ✅ RETURN ARRAY
        }
      } catch { }
    }
    return [];  // ✅ RETURN EMPTY ARRAY
  }
  async function fetchTerms(dclId, forDoc) {
    const flagField = forDoc === "CI" ? "cr650_is_printed_ci" : "cr650_is_printed_pl";
    const sel = ["cr650_termcondition", flagField, "_cr650_dcl_number_value"].join(",");
    const out = [];
    for (const f of [`_cr650_dcl_number_value eq guid('${dclId}')`, `_cr650_dcl_number_value eq '${dclId}'`]) {
      const url = `${API_URL_TERMS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}`;
      try {
        const r = await safeAjax(url);
        if (Array.isArray(r?.value)) {
          for (const row of r.value) { if (row[flagField] && row.cr650_termcondition) out.push(row.cr650_termcondition); }
          if (out.length) break;
        }
      } catch { }
    }
    return out;
  }

  async function fetchCharges(dclId) {
    const API_URL_CHARGES = "/_api/cr650_dcl_discounts_chargeses";

    // ✅ Step 1: Fetch without filter to see what's available
    const allUrl = `${API_URL_CHARGES}?$top=100`;

    try {
      const all = await safeAjax(allUrl);
      if (all && Array.isArray(all.value) && all.value.length > 0) {
        console.log("[fetchCharges] 🔍 Sample record to identify lookup field:", all.value[0]);
        console.log("[fetchCharges] 🔍 All field names:", Object.keys(all.value[0]));

        // ✅ Step 2: Filter client-side for now
        const normalizedDclId = (dclId || "").toLowerCase();
        const filtered = all.value.filter(row => {
          // Try different possible field names
          const ref1 = (row._cr650_dclreference_value || "").toLowerCase();
          const ref2 = (row._cr650_dcl_number_value || "").toLowerCase();
          const ref3 = (row["_cr650_dclreference_value@odata.bind"] || "").toLowerCase();

          return ref1 === normalizedDclId ||
            ref2 === normalizedDclId ||
            ref3.includes(normalizedDclId);
        });

        // ✅ Sort by creation date
        filtered.sort((a, b) => new Date(a.createdon || 0) - new Date(b.createdon || 0));

        const charges = filtered.map(row => {
          let displayName = row.cr650_name || "Unnamed Charge";
          if (displayName === "Other Charges" && row.cr650_other_charges_name) {
            displayName = row.cr650_other_charges_name;
          }
          return {
            id: row.cr650_dcl_discounts_chargesesid || "",
            name: displayName,
            amount: Number(row.cr650_totalimpact) || 0,
            currency: row.cr650_currency || "SAR",
            qty: row.cr650_qty || "",
            createdon: row.createdon || ""
          };
        });

        console.log(`[fetchCharges] Found ${charges.length} charges:`, charges);
        return charges;
      }
    } catch (err) {
      console.error("[fetchCharges] Error:", err);
    }

    return [];
  }

  async function fetchAdditionalDetails(dclId) {
    const API_URL_ADDITIONAL = "/_api/cr650_dcl_additional_detailses";
    const url = `${API_URL_ADDITIONAL}?$filter=_cr650_dclreference_value eq ${dclId}`;
    try {
      const r = await safeAjax(url);
      if (r?.value?.length > 0) {
        console.log("✅ Fetched Additional Details:", r.value[0]);
        return r.value[0];
      }
    } catch (err) {
      console.warn("Could not fetch additional details:", err);
    }
    return null;
  }
  async function fetchNotifyParties(dclId) {
    const sel = "cr650_notify_party,_cr650_dcl_number_value";
    for (const f of [`_cr650_dcl_number_value eq guid('${dclId}')`, `_cr650_dcl_number_value eq '${dclId}'`]) {
      const url = `${API_URL_NOTIFY}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}`;
      try {
        const r = await safeAjax(url);
        if (Array.isArray(r?.value)) {
          return r.value.map(row => row.cr650_notify_party).filter(Boolean);
        }
      } catch { }
    }
    return [];
  }

  async function fetchCustomerModels(dclId) {
    const API_URL_CUST_MODELS = "/_api/cr650_dcl_customer_models";
    const sel = "cr650_model_name,cr650_info,_cr650_dcl_id_value,cr650_dcl_customer_modelid";
    const normalizedDclId = (dclId || "").toLowerCase();

    // Try server-side filter first
    for (const f of [`_cr650_dcl_id_value eq ${dclId}`, `_cr650_dcl_id_value eq '${dclId}'`]) {
      try {
        const r = await safeAjax(`${API_URL_CUST_MODELS}?$select=${encodeURIComponent(sel)}&$filter=${encodeURIComponent(f)}&$orderby=createdon asc`);
        if (Array.isArray(r?.value)) {
          console.log(`[fetchCustomerModels] Found ${r.value.length} customer models for DCL ${dclId}`);
          return r.value;
        }
      } catch { /* try next filter */ }
    }

    // Fallback: fetch all and filter client-side
    try {
      const all = await safeAjax(`${API_URL_CUST_MODELS}?$top=500`);
      if (all && Array.isArray(all.value)) {
        const filtered = all.value.filter(row =>
          (row._cr650_dcl_id_value || "").toLowerCase() === normalizedDclId
        );
        console.log(`[fetchCustomerModels] Fallback found ${filtered.length} customer models`);
        return filtered;
      }
    } catch (err) {
      console.error("[fetchCustomerModels] Error:", err);
    }
    return [];
  }

  /* =========================
     AR REPORT FUNCTIONS (for Update All)
     ========================= */

  /**
   * Get customer number from DCL Master record
   */
  async function getCustomerNumberFromDcl(dclId) {
    const url = `${API_URL_MASTER}(${dclId})?$select=cr650_customernumber`;
    try {
      const r = await safeAjax(url);
      return r?.cr650_customernumber || null;
    } catch (err) {
      console.error("Failed to get customer number from DCL:", err);
      return null;
    }
  }

  /**
   * Fetch AR Report records by customer number
   */
  async function fetchARReportByCustomer(customerNumber) {
    if (!customerNumber) return [];

    const selectCols = "cr650_salesordernumber,cr650_itemno,cr650_qty,cr650_price,cr650_vatlineamount,cr650_trxlineamount";
    const filter = `cr650_customernumber eq '${customerNumber}'`;
    const url = `${AR_REPORT_API}?$select=${encodeURIComponent(selectCols)}&$filter=${encodeURIComponent(filter)}&$top=5000`;

    try {
      const r = await safeAjax(url);
      return r?.value || [];
    } catch (err) {
      console.error("Failed to fetch AR Report:", err);
      return [];
    }
  }

  /**
   * Build AR Index lookup map by order number and item code
   * Key: "{orderNumber}|{itemCode}"
   * Value: { qty, price, vat, lineAmount }
   */
  function buildARIndexByOrderAndItem(arReportData) {
    const index = new Map();

    console.log("📊 Building AR Index from", arReportData.length, "records");

    // Log first record to see the raw data structure
    if (arReportData.length > 0) {
      console.log("📊 Sample AR record (raw):", arReportData[0]);
    }

    for (const ar of arReportData || []) {
      const orderNo = String(ar.cr650_salesordernumber || "").trim();
      const itemCode = String(ar.cr650_itemno || "").trim();

      if (!orderNo || !itemCode) continue;

      const key = `${orderNo}|${itemCode}`;

      // Parse values - handle both string and number types
      const qty = parseFloat(ar.cr650_qty) || 0;
      const price = parseFloat(ar.cr650_price) || 0;
      const vat = parseFloat(ar.cr650_vatlineamount) || 0;
      const lineAmount = parseFloat(ar.cr650_trxlineamount) || 0;

      // Log parsed values for debugging
      console.log(`📊 AR Index: ${key} => qty:${qty}, price:${price}, vat:${vat}, lineAmt:${lineAmount}`);

      index.set(key, {
        qty,
        price,
        vat,
        lineAmount
      });
    }

    return index;
  }

  /**
   * Update a table row with AR Report data
   * Updates: Unit Price (cr650_price → cr650_unitprice) and 5% VAT (cr650_vatlineamount → cr650_vat5percent)
   */
  async function updateRowFromARData(tr, arData) {
    const loadingQtyInput = tr.querySelector(".loading-qty");
    const unitPriceInput = tr.querySelector(".unit-price");  // This is an <input>, not a cell
    const vatCell = tr.querySelector(".vat");  // This is a <td> cell

    let updated = false;

    console.log("📝 AR Data for row:", arData);

    // Update loading quantity if AR has qty > 0
    if (arData.qty > 0 && loadingQtyInput) {
      loadingQtyInput.value = fmt2(arData.qty);
      console.log("  ✓ Updated loading qty:", arData.qty);
      updated = true;
    }

    // Update unit price from AR Report (cr650_price → cr650_unitprice)
    // Unit price is an <input> element, so use .value
    if (unitPriceInput && arData.price != null) {
      const priceValue = Math.abs(arData.price);
      if (priceValue > 0) {
        unitPriceInput.value = fmt2(priceValue);
        console.log("  ✓ Updated unit price:", priceValue);
        updated = true;
      }
    }

    // Update 5% VAT from AR Report (cr650_vatlineamount → cr650_vat5percent)
    // VAT is a <td> cell, so use .textContent
    if (vatCell && arData.vat != null) {
      const vatValue = Math.abs(arData.vat);
      if (vatValue > 0) {
        vatCell.textContent = fmt2(vatValue);
        console.log("  ✓ Updated VAT:", vatValue);
        updated = true;
      }
    }

    if (updated) {
      // Recalculate derived values (including prices since AR data changed them)
      recalcRow(tr, { prices: true });
      markRowDirty(tr);
      return true;
    }

    return false;
  }

  /* =========================
     SHIPPED ORDERS FUNCTIONS (for Container # dropdown)
     ========================= */

  // Global shipped index cache
  let SHIPPED_INDEX_CACHE = new Map();

  /**
   * Fetch shipped orders by sales order number
   */
  async function fetchShippedBySo(orderNumber) {
    if (!orderNumber) return [];

    const selectCols = SHIPPED_FIELDS.join(",");
    const filter = `cr650_order_number eq '${orderNumber}'`;
    const orderby = "createdon desc";
    const url = `${SHIPPED_API}?$select=${encodeURIComponent(selectCols)}&$filter=${encodeURIComponent(filter)}&$orderby=${encodeURIComponent(orderby)}&$top=5000`;

    try {
      const r = await safeAjax(url);
      return r?.value || [];
    } catch (err) {
      console.warn("Failed to fetch shipped orders for", orderNumber, err);
      return [];
    }
  }

  /**
   * Build shipped index from raw shipped rows
   * Returns Map: "{orderNo}|{itemCode}" → Array of { dn, dateDisplay, whenTs }
   */
  function buildShippedIndex(rows) {
    const map = new Map();

    for (const r of rows || []) {
      const orderNo = String(r.cr650_order_number || "").trim();
      const itemCode = String(r.cr650_item_no || "").trim();

      if (!orderNo || !itemCode) continue;

      const key = `${orderNo}|${itemCode}`;

      // Extract container no + seal no
      const rawContainer = r.cr650_container_no != null ? String(r.cr650_container_no).trim() : "";
      if (!rawContainer) continue;

      const sealNo = r.cr650_seal_no != null ? String(r.cr650_seal_no).trim() : "";
      // Take part before "/" and append "-sealNo"
      const containerBase = rawContainer.indexOf("/") !== -1 ? rawContainer.substring(0, rawContainer.indexOf("/")) : rawContainer;
      const dn = sealNo ? `${containerBase}-${sealNo}` : containerBase;

      // Parse date
      const shipDateStr = r.cr650_shipment_date || "";
      const createdOn = r.createdon || "";
      const whenTs = Date.parse(shipDateStr) || Date.parse(createdOn) || 0;

      // Format date for display
      let dateDisplay = "";
      if (shipDateStr) {
        dateDisplay = shipDateStr;
      } else if (createdOn) {
        const d = new Date(createdOn);
        dateDisplay = d.toLocaleDateString();
      }

      // Get or create array for this key
      const arr = map.get(key) || [];
      arr.push({
        dn,
        dateDisplay,
        whenTs
      });
      map.set(key, arr);
    }

    // Sort each array by date (most recent first)
    for (const [k, arr] of map.entries()) {
      arr.sort((a, b) => b.whenTs - a.whenTs);
    }

    return map;
  }

  /**
   * Get unique DN options for an order/item combination
   * Returns array of { value, label } for dropdown
   */
  function getContainerOptionsForItem(orderNo, itemCode) {
    const key = `${orderNo}|${itemCode}`;
    const shipped = SHIPPED_INDEX_CACHE.get(key) || [];

    // Get unique DN numbers with their dates
    const dnMap = new Map();
    for (const s of shipped) {
      if (s.dn && !dnMap.has(s.dn)) {
        dnMap.set(s.dn, {
          value: s.dn,
          label: `${s.dn} (${s.dateDisplay})`,
          dateDisplay: s.dateDisplay
        });
      }
    }

    return Array.from(dnMap.values());
  }

  /**
   * Fetch shipped data for all unique order numbers in LP rows
   */
  async function fetchShippedForAllOrders(lpRows) {
    const orderNumbers = new Set();

    for (const row of lpRows || []) {
      const orderNo = row.cr650_ordernumber || row.orderNo;
      if (orderNo) orderNumbers.add(String(orderNo).trim());
    }

    console.log(`📦 Fetching shipped data for ${orderNumbers.size} order numbers...`);

    // Clear cache
    SHIPPED_INDEX_CACHE = new Map();

    // Fetch shipped data for ALL orders in parallel
    const results = await Promise.allSettled(
      Array.from(orderNumbers).map(async (orderNo) => {
        const shippedRows = await fetchShippedBySo(orderNo);
        return { orderNo, shippedRows };
      })
    );

    // Merge all results into cache
    for (const result of results) {
      if (result.status !== "fulfilled") continue;
      const shippedIndex = buildShippedIndex(result.value.shippedRows);
      for (const [key, arr] of shippedIndex.entries()) {
        const existing = SHIPPED_INDEX_CACHE.get(key) || [];
        SHIPPED_INDEX_CACHE.set(key, [...existing, ...arr]);
      }
    }

    console.log(`📦 Built shipped index with ${SHIPPED_INDEX_CACHE.size} order-item combinations`);
  }

  /**
   * Populate container dropdown with shipped options (DN number + date) for a row
   */
  function populateContainerDropdown(tr) {
    const orderNoCell = tr.querySelector(".order-no");
    const itemCodeCell = tr.querySelector(".item-code");
    const containerInput = tr.querySelector(".container-no");
    const datalistId = `container-options-${tr.rowIndex || Math.random().toString(36).substr(2, 9)}`;

    if (!containerInput || !orderNoCell || !itemCodeCell) return;

    const orderNo = (orderNoCell.textContent || "").trim();
    const itemCode = (itemCodeCell.textContent || "").trim();

    if (!orderNo || !itemCode) return;

    // Get DN options from shipped data (now returns { value, label, dateDisplay })
    const options = getContainerOptionsForItem(orderNo, itemCode);

    // Create or update datalist
    let datalist = tr.querySelector(`#${datalistId}`);
    if (!datalist) {
      datalist = document.createElement("datalist");
      datalist.id = datalistId;
      containerInput.after(datalist);
      containerInput.setAttribute("list", datalistId);
    }

    // Clear and populate options with DN number as value, label shows DN + date
    datalist.innerHTML = "";
    for (const opt of options) {
      const option = document.createElement("option");
      option.value = opt.value;  // DN number
      option.label = opt.label;  // "DN# (date)" - shows in dropdown
      option.textContent = opt.label;  // For browsers that use textContent
      datalist.appendChild(option);
    }

    if (options.length > 0) {
      console.log(`📦 Populated ${options.length} DN options for ${orderNo}|${itemCode}:`, options);
    }
  }

  /**
   * Populate container dropdowns for all rows
   */
  function populateAllContainerDropdowns() {
    const rows = QA("#itemsTableBody tr.lp-data-row");
    rows.forEach(populateContainerDropdown);
  }

  /* =========================
     HS CODE MAPPING FUNCTIONS
     ========================= */

  // Global HS codes cache - maps normalized item number to HS code
  let HS_CODES_CACHE = new Map();

  /**
   * Normalize item number by removing commas and trimming
   * Handles both "000,000,000,000" and "000000000000" formats
   */
  function normalizeItemNumber(itemNumber) {
    if (!itemNumber) return "";
    return String(itemNumber).replace(/,/g, "").trim();
  }

  /**
   * Fetch all HS codes from Dataverse and build cache
   */
  async function fetchAndCacheHSCodes() {
    const selectCols = HS_CODES_FIELDS.join(",");
    const url = `${HS_CODES_API}?$select=${encodeURIComponent(selectCols)}&$top=5000`;

    try {
      console.log("🏷️ Fetching HS codes...");
      const r = await safeAjax(url);
      const rows = r?.value || [];

      // Clear cache and rebuild
      HS_CODES_CACHE = new Map();

      for (const row of rows) {
        const itemNumber = row.cr650_itemnumber;
        const hsCode = row.cr650_hscode;

        if (!itemNumber || !hsCode) continue;

        // Normalize the item number (remove commas) for lookup
        const normalizedItemNo = normalizeItemNumber(itemNumber);

        // Store both the normalized version and original (with commas) for lookup
        HS_CODES_CACHE.set(normalizedItemNo, {
          hsCode: hsCode,
          itemDescription: row.cr650_itemdescription || "",
          category: row.cr650_category || ""
        });

        // Also store the original format in case someone searches with commas
        if (itemNumber !== normalizedItemNo) {
          HS_CODES_CACHE.set(itemNumber, {
            hsCode: hsCode,
            itemDescription: row.cr650_itemdescription || "",
            category: row.cr650_category || ""
          });
        }
      }

      console.log(`🏷️ Cached ${HS_CODES_CACHE.size} HS code mappings`);
      return HS_CODES_CACHE;
    } catch (err) {
      console.error("Failed to fetch HS codes:", err);
      return new Map();
    }
  }

  /**
   * Look up HS code for an item number
   * Handles both comma and no-comma formats
   */
  function lookupHSCode(itemNumber) {
    if (!itemNumber) return null;

    // Try exact match first
    let result = HS_CODES_CACHE.get(itemNumber);
    if (result) return result;

    // Try normalized (no commas) version
    const normalized = normalizeItemNumber(itemNumber);
    result = HS_CODES_CACHE.get(normalized);
    if (result) return result;

    // Not found
    return null;
  }

  /**
   * Auto-populate HS code for a single row based on item code
   */
  function populateHSCodeForRow(tr) {
    const itemCodeCell = tr.querySelector(".item-code");
    const hsCodeInput = tr.querySelector(".hs-code");

    if (!itemCodeCell || !hsCodeInput) return false;

    const itemCode = (itemCodeCell.textContent || "").trim();
    if (!itemCode) return false;

    // Look up HS code
    const hsData = lookupHSCode(itemCode);

    if (hsData && hsData.hsCode) {
      // Remove commas from HS code for storage/display
      const hsCodeClean = String(hsData.hsCode).replace(/,/g, "").trim();

      // Only update if empty or different
      const currentHsCode = (hsCodeInput.value || "").replace(/,/g, "").trim();
      if (!currentHsCode || currentHsCode !== hsCodeClean) {
        hsCodeInput.value = hsCodeClean;
        console.log(`🏷️ Set HS code for ${itemCode}: ${hsCodeClean}`);
        return true;
      }
    }

    return false;
  }

  /**
   * Auto-populate HS codes for all rows in the table
   */
  function populateAllHSCodes() {
    const rows = QA("#itemsTableBody tr.lp-data-row");
    let updatedCount = 0;

    rows.forEach(tr => {
      if (populateHSCodeForRow(tr)) {
        updatedCount++;
      }
    });

    if (updatedCount > 0) {
      console.log(`🏷️ Auto-populated ${updatedCount} HS codes`);
    }
  }

  /**
   * Handle item code change - update HS code automatically
   */
  function onItemCodeChange(tr) {
    if (!tr) return;

    // Small delay to ensure the cell content is updated
    setTimeout(() => {
      if (populateHSCodeForRow(tr)) {
        markRowDirty(tr);
      }
    }, 100);
  }

  /* =========================
   5) HEADER / FOOTER (shared)
   ========================= */

  async function buildHeader(master, brandsText, docType = "", orderNumbers = [], notifyParties = [], containers = [], lang = "EN", hsCodes = [], customerModels = []) {
    const {
      Header, Paragraph, TextRun, Table, TableRow, TableCell,
      WidthType, BorderStyle, AlignmentType, VerticalAlign,
      ImageRun, TabStopType
    } = w.docx;

    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    const companyType = getCompanyType(master);
    const LINE_1_5 = 360;
    const none = { style: BorderStyle.NONE };
    const noBorders = { top: none, bottom: none, left: none, right: none };
    // ✅ Extract unique container types
    // ✅ Extract unique container types and map to labels
    const containerTypes = [...new Set(
      containers
        .map(c => c[CF.type])
        .filter(Boolean)
        .map(value => getContainerTypeLabel(value))  // ✅ Convert value to label
        .filter(Boolean)  // ✅ Remove any empty strings
    )].join(", ");

    const p = (children, align = "LEFT") =>
      new Paragraph({
        alignment: AlignmentType[align],
        bidirectional: rtl,
        spacing: { line: LINE_1_5, lineRule: "AUTO" },
        children
      });

    const ucFirst = s => { const str = String(s ?? ""); return str.charAt(0).toUpperCase() + str.slice(1); };
    const label = t => new TextRun({ text: t ? ucFirst(t) + ": " : "", bold: true, size: 18, font: "Arial" });
    const val = v => new TextRun({ text: v ? ucFirst(String(v)) : "", size: 18, font: "Arial" });
    const text = t => new TextRun({ text: t ? ucFirst(String(t)) : "", size: 18, font: "Arial" });
    const soften = s =>
      ucFirst(String(s ?? ""))
        .replace(/,/g, ",\u200B")
        .replace(/\//g, "/\u200B")
        .replace(/-/g, "-\u200B");

    // Logos
    let logosLine;
    if (companyType === "TECHNOLUBE") {
      const techLogoBuf = await fetchArrayBuffer(TECHNOLUBE_LOGO_URL);
      logosLine = new Paragraph({
        alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
        spacing: { line: LINE_1_5, lineRule: "AUTO" },  // ✅ Added
        children: [
          ...(techLogoBuf
            ? [new ImageRun({ data: techLogoBuf, transformation: { width: 150, height: 67.5 } })]
            : [new TextRun({ text: "Technolube", bold: true, size: 24, font: "Arial" })])
        ]
      });
    } else {
      const arLogoBuf = await fetchArrayBuffer(AR_LOGO_URL);
      const enLogoBuf = await fetchArrayBuffer(EN_LOGO_URL);
      const leftLogo = rtl ? arLogoBuf : enLogoBuf;
      const rightLogo = rtl ? enLogoBuf : arLogoBuf;

      logosLine = new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: LINE_1_5, lineRule: "AUTO" },  // ✅ Added
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          ...(leftLogo
            ? [new ImageRun({ data: leftLogo, transformation: { width: 150, height: 67.5 } })]
            : []),
          new TextRun({ text: "\t" }),
          ...(rightLogo
            ? [new ImageRun({ data: rightLogo, transformation: { width: 150, height: 67.5 } })]
            : [])
        ]
      });
    }

    // Document Title
    const docTitle = new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200, line: LINE_1_5, lineRule: "AUTO" },  // ✅ Added line spacing
      children: [
        new TextRun({
          text: docType || "Commercial Invoice",
          bold: true,
          size: 28,
          font: "Arial",
          underline: {}
        })
      ]
    });

    const benefText = companyType === "TECHNOLUBE"
      ? soften("Techno Lube L.L.C. Techno Park, PO Box-116636, Jebel Ali, Dubai, UAE, Phone No.: +(971) 4 801 8444, Fax: +(971) 4 886 7014 VAT No.: 100206850800003")
      : soften("Petrolube Oil Company, Closed Joint Stock Company (Single Shareholder) / The CR number: 4030608930 / Company's Capital: 300,000,000 SR, Tel: +966126996600 / Address: Aya Mall, Prince Sultan Street / Al-Muhammadiyah District / Jeddah 23618, P.O. Box 41728");

    const party = (master[MF.party] || "").trim().toLowerCase();

    // Row 1 - Invoice details
    const row1Cells = [
      new TableCell({
        borders: noBorders,
        children: [new Paragraph({
          bidirectional: rtl,
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          spacing: { line: LINE_1_5, lineRule: "AUTO" },
          children: [label(L.invoiceNum), val(master[MF.ciNumber] || master[MF.dcl])]
        })]
      }),
      new TableCell({
        borders: noBorders,
        children: [
          new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [
              label(L.invoiceDate),
              val(
                new Date().toLocaleDateString("en-GB", {
                  weekday: "long",
                  day: "numeric",
                  month: "long",
                  year: "numeric"
                })
              )
            ]
          })
        ]
      })
    ];

    const row1 = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
      columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
      rows: [
        new TableRow({
          children: rtl ? row1Cells.reverse() : row1Cells
        })
      ]
    });

    // Row 1b - L/C Number and L/C Issue Date (only if L/C Number exists)
    let rowLC = null;
    const lcNumber = master[MF.lcNumber];
    if (lcNumber) {
      const lcIssueDateRaw = master[MF.lcIssueDate];
      const lcIssueDateFormatted = lcIssueDateRaw
        ? new Date(lcIssueDateRaw).toLocaleDateString("en-GB", {
            weekday: "long",
            day: "numeric",
            month: "long",
            year: "numeric"
          })
        : "";

      const rowLCCells = [
        new TableCell({
          borders: noBorders,
          children: [new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [label(L.lcNumber), val(lcNumber)]
          })]
        }),
        new TableCell({
          borders: noBorders,
          children: [new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [label(L.lcIssueDate), val(lcIssueDateFormatted)]
          })]
        })
      ];

      rowLC = new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
        columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
        rows: [
          new TableRow({
            children: rtl ? rowLCCells.reverse() : rowLCCells
          })
        ]
      });
    }

    // Row 2 - Beneficiary + Customer Models
    const rightColChildren = [];

    if (customerModels && customerModels.length > 0) {
      customerModels.forEach(cm => {
        const modelName = (cm.cr650_model_name || "").trim();
        const modelInfo = (cm.cr650_info || "").trim();
        if (modelName) {
          rightColChildren.push(
            new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { line: LINE_1_5, lineRule: "AUTO" },
              children: [label(modelName)]
            })
          );
        }
        if (modelInfo) {
          rightColChildren.push(
            new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { line: LINE_1_5, lineRule: "AUTO" },
              children: [text(soften(modelInfo))]
            })
          );
        }
      });
    }

    // If no customer models, leave right column empty
    if (rightColChildren.length === 0) {
      rightColChildren.push(
        new Paragraph({
          children: [new TextRun({ text: "", size: 18, font: "Arial" })]
        })
      );
    }

    const row2Cells = [
      new TableCell({
        width: { size: CONTENT_W / 2, type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        borders: noBorders,
        children: [
          new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [label(L.beneficiary)]
          }),
          new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [text(benefText)]
          })
        ]
      }),
      new TableCell({
        width: { size: CONTENT_W / 2, type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        borders: noBorders,
        children: rightColChildren
      })
    ];

    const row2 = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
      columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
      rows: [
        new TableRow({
          children: rtl ? row2Cells.reverse() : row2Cells
        })
      ]
    });

    // Notify Party Rows
    let notifyPartyTable = null;
    if (notifyParties && notifyParties.length > 0) {
      const isMultiple = notifyParties.length > 1;

      const notifyRows = notifyParties.map((party, index) => {
        // Build the label based on single vs multiple notify parties
        const notifyLabel = isMultiple
          ? `${L.notifyParty} #${index + 1}: `
          : `${L.notifyParty}: `;

        const fullWidthCell = new TableCell({
          width: { size: CONTENT_W, type: WidthType.DXA },
          borders: noBorders,
          children: [
            new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { line: LINE_1_5, lineRule: "AUTO" },
              children: [
                new TextRun({ text: notifyLabel, bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: party || "", size: 18, font: "Arial" })
              ]
            })
          ]
        });

        return new TableRow({
          children: [fullWidthCell]
        });
      });

      notifyPartyTable = new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
        columnWidths: [CONTENT_W],
        rows: notifyRows
      });
    }

    // Grid section - remaining fields (skip rows where both values are empty)
    const allPairs = [
      [L.shippingMarks, (brandsText && brandsText.trim()) ? brandsText.trim() : "", L.incoterms, master[MF.incoterms] || ""],
      [L.typeOfTransport, containerTypes || master[MF.transportMode] || "", L.countryOfOrigin, master[MF.countryOfOrigin] || ""],
      [L.transportMode, master[MF.transportMode] || "", L.currency, master[MF.currency] || ""],
      [L.portOfLoading, master[MF.loadingPort] || "", L.portOfDischarge, master[MF.destinationPort] || ""],
      [L.piNumber, master[MF.piNumber] || "", L.paymentTerms, master[MF.paymentTerms] || ""],
      [L.customerPoNumber, master[MF.customerPoNumber] || "", L.orderNumbers, (orderNumbers && orderNumbers.length > 0) ? orderNumbers.join(", ") : ""]
    ];

    // Filter: skip entire row if both left and right values are empty
    const pairs = allPairs.filter(([, v1, , v2]) => v1 || v2);

    const emptyCell = () => new TableCell({
      borders: noBorders,
      children: [new Paragraph({ children: [] })]
    });

    const gridRows = pairs.map(([l1, v1, l2, v2]) => {
      const leftCell = v1
        ? new TableCell({
            borders: noBorders,
            children: [new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { line: LINE_1_5, lineRule: "AUTO" },
              children: [label(l1), val(v1)]
            })]
          })
        : emptyCell();

      const rightCell = v2
        ? new TableCell({
            borders: noBorders,
            children: [new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { line: LINE_1_5, lineRule: "AUTO" },
              children: [label(l2), val(v2)]
            })]
          })
        : emptyCell();

      return new TableRow({
        children: rtl ? [rightCell, leftCell] : [leftCell, rightCell]
      });
    });

    const grid = pairs.length > 0
      ? new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
          columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
          rows: gridRows
        })
      : null;

    const headerChildren = [logosLine, docTitle, row1];
    if (rowLC) {
      headerChildren.push(rowLC);
    }
    headerChildren.push(row2);

    if (notifyPartyTable) {
      headerChildren.push(notifyPartyTable);
    }

    if (grid) {
      headerChildren.push(grid);
    }

    // Description of Goods and/or Services (two-column row, only if not null)
    const descGoods = (master[MF.descriptionGoods] || "").trim();
    if (descGoods) {
      const descCells = [
        new TableCell({
          borders: noBorders,
          children: [new Paragraph({
            bidirectional: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: LINE_1_5, lineRule: "AUTO" },
            children: [label(L.descriptionGoods), val(descGoods)]
          })]
        }),
        new TableCell({
          borders: noBorders,
          children: [new Paragraph({ children: [] })]
        })
      ];
      const descGoodsRow = new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
        columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
        rows: [
          new TableRow({
            children: rtl ? descCells.reverse() : descCells
          })
        ]
      });
      headerChildren.push(descGoodsRow);
    }

    // Add HS Code row for Packing List only (after grid, as last item)
    if (docType === "Packing List" && hsCodes && hsCodes.length > 0) {
      const hsCodesText = hsCodes.join(", ");
      const hsCodeRow = new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { ...noBorders, insideHorizontal: none, insideVertical: none },
        columnWidths: [CONTENT_W],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders,
                children: [new Paragraph({
                  bidirectional: rtl,
                  alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
                  spacing: { line: LINE_1_5, lineRule: "AUTO" },
                  children: [label(L.hsCode + ": "), val(hsCodesText)]
                })]
              })
            ]
          })
        ]
      });
      headerChildren.push(hsCodeRow);
    }

    return new Header({ children: headerChildren });
  }

  async function buildFooter(masterRec) {
    const {
      Footer, Paragraph, TextRun, AlignmentType, ImageRun, BorderStyle,
      PageNumber  // ✅ ADD THIS
    } = w.docx;
    const size7 = 14;
    const companyType = getCompanyType(masterRec);

    if (companyType === "TECHNOLUBE") {
      const arText = "ص.ب : ١١٦٦٣٦، مجمع الصناعات الوطنية، دبي - الامارات العربية المتحدة، هاتف : ٩٧١٤٨٠١٨٤٤٤+، فاكس : ٩٧١٤٨٨٦٧٠١٤+";
      const enText = "P.O. Box: 116636, National Industries Park, Dubai, U.A.E., Tel: +971 4 8018444, Fax: +971 4 8867014";
      const trnText = "Techno Lube Group TRN: 100206850800003";

      let footerLogoBuf = null;
      try { footerLogoBuf = await fetchArrayBuffer(TECHNOLUBE_FOOTER_LOGO_URL); } catch { }

      const children = [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 40 },
          children: [
            new TextRun({ text: arText, size: 12, font: "Arial", rightToLeft: true })
          ]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 40 },
          children: [
            new TextRun({ text: enText, size: 12, font: "Arial" })
          ]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 40, after: 80 },
          children: [
            new TextRun({ text: trnText, size: 12, font: "Arial", bold: true })
          ]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" }
          },
          spacing: { before: 0, after: 80 },
          children: []
        })
      ];

      if (footerLogoBuf) {
        children.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 0 },
          children: [
            new ImageRun({
              data: footerLogoBuf,
              transformation: { width: 450, height: 45 }
            })
          ]
        }));
      }

      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 0 },
        children: [
          new TextRun({ text: "This document was automatically generated by the system", font: "Arial", size: 14, italics: true, color: "666666" })
        ]
      }));

      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 0 },
        children: [
          new TextRun({
            children: [
              "Page ",
              PageNumber.CURRENT,
              " of ",
              PageNumber.TOTAL_PAGES
            ],
            font: "Arial",
            size: 16
          })
        ]
      }));

      return new Footer({ children });
    } else {
      // Petrolube Footer
      const en = "Petrolube Oil Company, Closed Joint Stock Company (Single Shareholder) / The CR number: 4030608830 / Company's Capital: 300,000,000 SR, Tel: +966126996600 / Address: Aya Mall, Prince Sultan Street / Al-Muhammadiyah District / Jeddah 23618, P.O. Box 41728";
      const ar = "شركة بترولوب للزيوت، شركة مساهمة مقفلة (شخص واحد) / رقم السجل التجاري: ٤٠٣٠٦٠٨٨٣٠ / رأس مال الشركة: ٣٠٠٬٠٠٠٬٠٠٠ ريال سعودي\nالهاتف: ‎+٩٦٦١٢٦٩٩٦٦٠٠ / العنوان: آية مول، شارع الأمير سلطان / حي المحمدية / جدة ٢٣٦١٨، ص.ب ٤١٧٢٨";

      return new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: en, size: size7, font: "Arial" })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: ar, size: size7, font: "Arial" })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 80, after: 0 },
            children: [
              new TextRun({ text: "This document was automatically generated by the system", font: "Arial", size: 14, italics: true, color: "666666" })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 0 },
            children: [
              new TextRun({
                children: [
                  "Page ",
                  PageNumber.CURRENT,
                  " of ",
                  PageNumber.TOTAL_PAGES
                ],
                font: "Arial",
                size: 16
              })
            ]
          })
        ]
      });
    }
  }

  function buildAdditionalDetailsTable(additionalDetails, docType, lang = "EN") {
    if (!additionalDetails) return null;

    const {
      Table, TableRow, TableCell, Paragraph, TextRun,
      WidthType, BorderStyle, VerticalAlign, AlignmentType, ShadingType
    } = w.docx;
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;

    const printOn = additionalDetails.cr650_printon;
    const shouldPrint =
      printOn === 3 ||
      (docType === "CI" && printOn === 1) ||
      (docType === "PL" && printOn === 2);

    if (!shouldPrint) {
      console.log(`ℹ️ Additional details not printed on ${docType} (cr650_printon=${printOn})`);
      return null;
    }

    const hiddenStr = additionalDetails.cr650_hidden_details || "";
    const hiddenSet = new Set(hiddenStr.split(",").map(x => x.trim()).filter(Boolean));

    const allDetails = [
      {
        key: "totalPallets",
        label: "Number of pallets the products are packed on",
        field: "cr650_totalpallets",
        format: (v) => fmtNum(v, 2)
      },
      {
        key: "totalCartons",
        label: "Total Number of Cartons",
        field: "cr650_totalcartons",
        format: (v) => String(Math.round(v) || 0)
      },
      {
        key: "totalDrums",
        label: "Total Number of Drums",
        field: "cr650_totaldrums",
        format: (v) => String(Math.round(v) || 0)
      },
      {
        key: "totalPails",
        label: "Total Number of Pails",
        field: "cr650_totalpails",
        format: (v) => String(Math.round(v) || 0)
      },
      {
        key: "totalPackages",
        label: "Total Number of Packages",
        field: "cr650_totalpackages",
        format: (v) => String(Math.round(v) || 0)
      },
      {
        key: "totalNetWeight",
        label: "Total Net Weight of Packages (Kgs)",
        field: "cr650_totalnetweight",
        format: (v) => fmtNum(v, 2)
      },
      {
        key: "totalGrossWeight",
        label: "Total Gross Weight of Packages (Kgs)",
        field: "cr650_totalgrossweight",
        format: (v) => fmtNum(v, 2)
      }
    ];

    const visibleDetails = allDetails.filter(d => !hiddenSet.has(d.key));
    if (!visibleDetails.length) return null;

    const GRID = { style: BorderStyle.SINGLE, size: 8, color: "000000" };
    const OUTER = { style: BorderStyle.SINGLE, size: 12, color: "000000" };
    const rows = [];

    visibleDetails.forEach(detail => {
      const rawValue = additionalDetails[detail.field];
      const displayValue = detail.format ? detail.format(rawValue) : String(rawValue ?? "0");

      rows.push(new TableRow({
        children: [
          new TableCell({
            width: { size: Math.floor(CONTENT_W * 0.7), type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [new TextRun({ text: detail.label, font: "Arial", size: 18 })]
              })
            ]
          }),
          new TableCell({
            width: { size: Math.floor(CONTENT_W * 0.3), type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: displayValue, font: "Arial", size: 18 })]
              })
            ]
          })
        ]
      }));
    });

    const comments = additionalDetails.cr650_additionalcomments;
    if (comments && comments.trim()) {
      rows.push(new TableRow({
        children: [
          new TableCell({
            width: { size: Math.floor(CONTENT_W * 0.7), type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                children: [new TextRun({ text: L.additionalCommentsLabel, font: "Arial", size: 18 })]
              })
            ]
          }),
          new TableCell({
            width: { size: Math.floor(CONTENT_W * 0.3), type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                children: [new TextRun({ text: comments.trim(), font: "Arial", size: 18 })]
              })
            ]
          })
        ]
      }));
    }

    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [
        Math.floor(CONTENT_W * 0.7),
        Math.floor(CONTENT_W * 0.3)
      ],
      borders: {
        top: OUTER,
        bottom: OUTER,
        left: OUTER,
        right: OUTER,
        insideHorizontal: GRID,
        insideVertical: GRID
      },
      rows
    });
  }

  /* =========================
     6) CI BUILDERS
     ========================= */
  function buildCombinedTable_CI(rows, totals, lang = "EN", currency = "SAR", showHsCode = true) {
    const {
      Table, TableRow, TableCell, Paragraph, TextRun,
      WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign,
      TabStopType
    } = w.docx;
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    const GRID = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
    const OUTER = { style: BorderStyle.SINGLE, size: 10, color: "000000" };

    // Check if all VAT values are zero - if so, hide VAT columns
    const showVatColumns = rows.some(r => (Number(r.vat) || 0) !== 0);

    // Column widths for items section - adjust based on showVatColumns and showHsCode
    let CW;
    if (showVatColumns && showHsCode) {
      CW = { sn: 700, hs: 1700, item: 1900, desc: 3800, qty: 1200, unit: 1400, totalExVat: 1400, vat: 1200, total: 1800 };
    } else if (showVatColumns && !showHsCode) {
      CW = { sn: 700, item: 2200, desc: 4300, qty: 1200, unit: 1400, totalExVat: 1400, vat: 1200, total: 1800 };
    } else if (!showVatColumns && showHsCode) {
      CW = { sn: 700, hs: 1700, item: 1900, desc: 4400, qty: 1400, unit: 1600, total: 2100 };
    } else {
      CW = { sn: 700, item: 2200, desc: 5100, qty: 1400, unit: 1600, total: 2100 };
    }
    const sum = Object.values(CW).reduce((a, b) => a + b, 0);
    const scale = CONTENT_W / sum;
    Object.keys(CW).forEach(k => CW[k] = Math.floor(CW[k] * scale));

    const th = (t, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, color: "auto", fill: HEADER_BG },
      borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
      children: [
        new Paragraph({
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.CENTER,
          bidirectional: rtl,
          children: [new TextRun({ text: t, bold: true, font: "Arial", size: 18 })]
        })
      ]
    });

    const td = (t, w, align = "C", bold = false) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
      children: [
        new Paragraph({
          alignment:
            align === "R"
              ? AlignmentType.RIGHT
              : align === "L"
                ? (rtl ? AlignmentType.RIGHT : AlignmentType.LEFT)
                : (rtl ? AlignmentType.RIGHT : AlignmentType.CENTER),
          bidirectional: rtl,
          children: [new TextRun({ text: String(t ?? ""), font: "Arial", size: 18, bold })]
        })
      ]
    });

    // Items header - conditionally include VAT and HS Code columns
    let headerCells;
    if (showVatColumns && showHsCode) {
      headerCells = [
        th(L.sn, CW.sn), th(L.hsCode, CW.hs), th(L.itemCode, CW.item), th(L.itemDesc, CW.desc),
        th(L.qty, CW.qty), th(L.unitPrice, CW.unit), th(L.totalExVat, CW.totalExVat),
        th(L.vat5, CW.vat), th(L.totalPrice, CW.total)
      ];
    } else if (showVatColumns && !showHsCode) {
      headerCells = [
        th(L.sn, CW.sn), th(L.itemCode, CW.item), th(L.itemDesc, CW.desc),
        th(L.qty, CW.qty), th(L.unitPrice, CW.unit), th(L.totalExVat, CW.totalExVat),
        th(L.vat5, CW.vat), th(L.totalPrice, CW.total)
      ];
    } else if (!showVatColumns && showHsCode) {
      headerCells = [
        th(L.sn, CW.sn), th(L.hsCode, CW.hs), th(L.itemCode, CW.item), th(L.itemDesc, CW.desc),
        th(L.qty, CW.qty), th(L.unitPrice, CW.unit), th(L.totalPrice, CW.total)
      ];
    } else {
      headerCells = [
        th(L.sn, CW.sn), th(L.itemCode, CW.item), th(L.itemDesc, CW.desc),
        th(L.qty, CW.qty), th(L.unitPrice, CW.unit), th(L.totalPrice, CW.total)
      ];
    }

    const header = new TableRow({
      tableHeader: true,
      children: rtl ? headerCells.reverse() : headerCells
    });

    // Items body rows - conditionally include VAT and HS Code columns
    const itemRows = rows.map((r, i) => {
      let rowCells;
      if (showVatColumns && showHsCode) {
        rowCells = [
          td(i + 1, CW.sn, "C"), td(r.hsCode, CW.hs, "C"), td(r.itemCode, CW.item, "C"),
          td(r.description, CW.desc, "L"), td(fmtNum(r.qty, 0), CW.qty, "C"),
          td(fmtNum(r.unit, 2), CW.unit, "C"), td(fmtNum(r.totalExVat, 2), CW.totalExVat, "C"),
          td(fmtNum(r.vat, 2), CW.vat, "C"), td(fmtNum(r.total, 2), CW.total, "C")
        ];
      } else if (showVatColumns && !showHsCode) {
        rowCells = [
          td(i + 1, CW.sn, "C"), td(r.itemCode, CW.item, "C"), td(r.description, CW.desc, "L"),
          td(fmtNum(r.qty, 0), CW.qty, "C"), td(fmtNum(r.unit, 2), CW.unit, "C"),
          td(fmtNum(r.totalExVat, 2), CW.totalExVat, "C"), td(fmtNum(r.vat, 2), CW.vat, "C"),
          td(fmtNum(r.total, 2), CW.total, "C")
        ];
      } else if (!showVatColumns && showHsCode) {
        rowCells = [
          td(i + 1, CW.sn, "C"), td(r.hsCode, CW.hs, "C"), td(r.itemCode, CW.item, "C"),
          td(r.description, CW.desc, "L"), td(fmtNum(r.qty, 0), CW.qty, "C"),
          td(fmtNum(r.unit, 2), CW.unit, "C"), td(fmtNum(r.total, 2), CW.total, "C")
        ];
      } else {
        rowCells = [
          td(i + 1, CW.sn, "C"), td(r.itemCode, CW.item, "C"), td(r.description, CW.desc, "L"),
          td(fmtNum(r.qty, 0), CW.qty, "C"), td(fmtNum(r.unit, 2), CW.unit, "C"),
          td(fmtNum(r.total, 2), CW.total, "C")
        ];
      }

      return new TableRow({
        children: rtl ? rowCells.reverse() : rowCells
      });
    });

    // Helper to create currency + amount cell with tab stop
    function createAmountCell(currencyValue, amountValue, columnWidth, boldAmount = false) {
      return new TableCell({
        width: { size: columnWidth, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          tabStops: [
            {
              type: TabStopType.RIGHT,
              position: columnWidth - 100
            }
          ],
          children: currencyValue && amountValue ? [
            new TextRun({
              text: currencyValue,
              font: "Arial",
              size: 18,
              bold: boldAmount
            }),
            new TextRun({
              text: "\t" + fmtNum(amountValue, 2),
              font: "Arial",
              size: 18,
              bold: boldAmount
            })
          ] : []
        })]
      });
    }

    // Totals section helper function
    // targetColumn: "total" for Total Price column, "vat" for 5% VAT column
    function createTotalRow(label, qtyValue, currencyValue, amountValue, targetColumn = "total", boldLabel = false, boldAmount = false, fullSpan = false) {

      // Calculate span count and width based on VAT and HS Code visibility
      let fullSpanCount, fullSpanWidth, labelSpanCount, labelSpanWidth;
      if (showVatColumns && showHsCode) {
        fullSpanCount = 8;
        fullSpanWidth = CW.sn + CW.hs + CW.item + CW.desc + CW.qty + CW.unit + CW.totalExVat + CW.vat;
        labelSpanCount = 4;
        labelSpanWidth = CW.sn + CW.hs + CW.item + CW.desc;
      } else if (showVatColumns && !showHsCode) {
        fullSpanCount = 7;
        fullSpanWidth = CW.sn + CW.item + CW.desc + CW.qty + CW.unit + CW.totalExVat + CW.vat;
        labelSpanCount = 3;
        labelSpanWidth = CW.sn + CW.item + CW.desc;
      } else if (!showVatColumns && showHsCode) {
        fullSpanCount = 6;
        fullSpanWidth = CW.sn + CW.hs + CW.item + CW.desc + CW.qty + CW.unit;
        labelSpanCount = 4;
        labelSpanWidth = CW.sn + CW.hs + CW.item + CW.desc;
      } else {
        fullSpanCount = 5;
        fullSpanWidth = CW.sn + CW.item + CW.desc + CW.qty + CW.unit;
        labelSpanCount = 3;
        labelSpanWidth = CW.sn + CW.item + CW.desc;
      }

      // Full span mode for VAT Total and Grand Total
      if (fullSpan) {
        const labelCell = new TableCell({
          columnSpan: fullSpanCount,
          width: { size: fullSpanWidth, type: WidthType.DXA },
          borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: rtl,
            children: [new TextRun({
              text: label,
              bold: boldLabel,
              font: "Arial",
              size: 18
            })]
          })]
        });

        // Last column: Currency + Amount (Total Price column)
        const totalCell = createAmountCell(currencyValue, amountValue, CW.total, boldAmount);

        const cellsArray = rtl
          ? [totalCell, labelCell]
          : [labelCell, totalCell];

        return new TableRow({ children: cellsArray });
      }

      // Standard mode (non-full span)
      // Label cell (merged)
      const labelCell = new TableCell({
        columnSpan: labelSpanCount,
        width: { size: labelSpanWidth, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          bidirectional: rtl,
          children: [new TextRun({
            text: label,
            bold: boldLabel,
            font: "Arial",
            size: 18
          })]
        })]
      });

      // Column 5: Qty
      const qtyCell = new TableCell({
        width: { size: CW.qty, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: qtyValue ? String(qtyValue) : "",
            font: "Arial",
            size: 18
          })]
        })]
      });

      // Remaining columns depend on targetColumn and whether VAT columns are shown
      let middleCells = [];

      if (showVatColumns) {
        // With VAT columns: 9 total columns
        if (targetColumn === "vat") {
          // Empty cells for columns 6-7, amount in column 8 (VAT), empty column 9
          const emptyCell67 = new TableCell({
            columnSpan: 2,
            width: { size: CW.unit + CW.totalExVat, type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ children: [] })]
          });

          const vatCell = createAmountCell(currencyValue, amountValue, CW.vat, boldAmount);

          const emptyCell9 = new TableCell({
            width: { size: CW.total, type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ children: [] })]
          });

          middleCells = [emptyCell67, vatCell, emptyCell9];
        } else {
          // Empty cells for columns 6-8, amount in column 9 (Total Price)
          const emptyCell678 = new TableCell({
            columnSpan: 3,
            width: { size: CW.unit + CW.totalExVat + CW.vat, type: WidthType.DXA },
            borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ children: [] })]
          });

          const totalCell = createAmountCell(currencyValue, amountValue, CW.total, boldAmount);

          middleCells = [emptyCell678, totalCell];
        }
      } else {
        // Without VAT columns: 7 total columns (sn, hs, item, desc, qty, unit, total)
        // Empty cell for column 6 (unit), amount in column 7 (Total Price)
        const emptyCell6 = new TableCell({
          width: { size: CW.unit, type: WidthType.DXA },
          borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [] })]
        });

        const totalCell = createAmountCell(currencyValue, amountValue, CW.total, boldAmount);

        middleCells = [emptyCell6, totalCell];
      }

      const cellsArray = rtl
        ? [...middleCells.reverse(), qtyCell, labelCell]
        : [labelCell, qtyCell, ...middleCells];

      return new TableRow({ children: cellsArray });
    }

    // Totals rows
    const totalRows = [];

    // 1. Total Package (qty only, no amount)
    totalRows.push(createTotalRow(
      L.totalPackage,
      fmtNum(totals.totalPackages, 0),
      "",
      "",
      "total",
      true,
      false,
      false  // ✅ Not full span
    ));

    // 2. Total (Ex VAT) - currency + amount in Total Price column (only if showing VAT columns)
    if (showVatColumns) {
      totalRows.push(createTotalRow(
        L.totalExVat,
        "",
        currency,
        totals.subtotal,
        "total",
        true,
        false,
        false  // ✅ Not full span
      ));
    }

    // 3. Dynamic charges - qty + currency + amount in Total Price column
    if (Array.isArray(totals.charges) && totals.charges.length > 0) {
      totals.charges.forEach(charge => {
        totalRows.push(createTotalRow(
          charge.name,
          charge.qty || "",
          charge.currency,
          charge.amount,
          "total",
          true,
          false,
          false  // ✅ Not full span
        ));
      });
    }

    // 4. 5 % VAT Total - FULL SPAN (only if showing VAT columns)
    if (showVatColumns) {
      totalRows.push(createTotalRow(
        L.vat5Total,
        "",
        currency,
        totals.vatTotal,
        "total",
        true,
        false,
        true  // ✅ FULL SPAN - merge all except last column
      ));
    }

    // 5. Grand Total - FULL SPAN (columns merged)
    totalRows.push(createTotalRow(
      L.grandTotal,
      "",
      currency,
      totals.grandTotal,
      "total",
      true,
      true,
      true  // ✅ FULL SPAN - merge all except last column
    ));

    // Combine all rows
    const allRows = [header, ...itemRows, ...totalRows];

    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: {
        top: OUTER,
        bottom: OUTER,
        left: OUTER,
        right: OUTER,
        insideHorizontal: GRID,
        insideVertical: GRID
      },
      rows: allRows
    });
  }

  function buildTotalsTable_CI(totals, lang = "EN", currency = "SAR") {
    const {
      Table, TableRow, TableCell, Paragraph, TextRun,
      WidthType, BorderStyle, AlignmentType, VerticalAlign
    } = w.docx;
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    const GRID = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
    const OUTER = { style: BorderStyle.SINGLE, size: 10, color: "000000" };

    // Column widths for 4-column layout
    const CW = {
      label: Math.floor(CONTENT_W * 0.5),   // 50% - Label column
      qty: Math.floor(CONTENT_W * 0.15),    // 15% - Quantity column
      currency: Math.floor(CONTENT_W * 0.15), // 15% - Currency column  
      amount: Math.floor(CONTENT_W * 0.2)   // 20% - Amount column
    };

    const rows = [];

    // Helper function to create a row with 4 columns
    function createRow(label, qty, curr, amount, boldLabel = false, boldAmount = false) {
      const labelCell = new TableCell({
        width: { size: CW.label, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          bidirectional: rtl,
          children: [new TextRun({
            text: label,
            bold: boldLabel,
            font: "Arial",
            size: 18
          })]
        })]
      });

      const qtyCell = new TableCell({
        width: { size: CW.qty, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: qty ? String(qty) : "",
            font: "Arial",
            size: 18
          })]
        })]
      });

      const currencyCell = new TableCell({
        width: { size: CW.currency, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: curr || "",
            font: "Arial",
            size: 18
          })]
        })]
      });

      const amountCell = new TableCell({
        width: { size: CW.amount, type: WidthType.DXA },
        borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: fmtNum(amount, 2),
            bold: boldAmount,
            font: "Arial",
            size: 18
          })]
        })]
      });

      const cellsArray = rtl
        ? [amountCell, currencyCell, qtyCell, labelCell]
        : [labelCell, qtyCell, currencyCell, amountCell];

      return new TableRow({ children: cellsArray });
    }

    // 1. Total Package row
    rows.push(createRow(
      L.totalPackage,
      fmtNum(totals.totalPackages, 0),
      "",
      "",
      true,
      false
    ));

    // 2. Subtotal row
    rows.push(createRow(
      L.subtotal,
      "",
      currency,
      totals.subtotal,
      true,
      false
    ));

    // 3. Dynamic charges rows (ordered by createdon)
    if (Array.isArray(totals.charges) && totals.charges.length > 0) {
      totals.charges.forEach(charge => {
        rows.push(createRow(
          charge.name,
          charge.qty,
          charge.currency,
          charge.amount,
          true,
          false
        ));
      });
    }

    // 4. VAT Total row
    rows.push(createRow(
      L.vat5Total,
      "",
      currency,
      totals.vatTotal,
      true,
      false
    ));

    // 5. Grand Total row (bold amount)
    rows.push(createRow(
      L.grandTotal,
      "",
      currency,
      totals.grandTotal,
      true,
      true
    ));

    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CW.label, CW.qty, CW.currency, CW.amount],
      borders: {
        top: OUTER,
        bottom: OUTER,
        left: OUTER,
        right: OUTER,
        insideHorizontal: GRID,
        insideVertical: GRID
      },
      rows
    });
  }

  async function buildDocument_CI(master, items, plans, brandsText, termsList, additionalDetails, lang = "EN", orderNumbers = [], notifyParties = [], containers = [], showHsCode = true, customerModels = []) {
    await ensureDocx();
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    const rows = items.map(it => {
      const p = plans[it[IF.planLkp]] || {};
      const qty = Number(it[IF.qty]) || 0;
      const unit = Number(p[LF_CI.unitPrice]) || 0;
      const vat = Number(p[LF_CI.vat5]) || 0;
      const totalExVat = qty * unit;  // ✅ NEW: qty * unit price (excluding VAT)
      const total = totalExVat + vat;  // ✅ UPDATED: total including VAT

      return {
        hsCode: p[LF_CI.hsCode] || "",
        itemCode: p[LF_CI.itemCode] || "",
        description: p[LF_CI.description] || p[LF_CI.packaging] || "",
        qty: qty,
        unit: unit,
        totalExVat: totalExVat,  // ✅ ADD THIS
        vat: vat,
        total: total
      };
    });
    const subtotal = rows.reduce((s, r) => s + (Number(r.unit) || 0) * (Number(r.qty) || 0), 0);
    const vatTotal = rows.reduce((s, r) => s + (Number(r.vat) || 0), 0);
    const totalPackages = rows.reduce((s, r) => s + (Number(r.qty) || 0), 0);
    const charges = await fetchCharges(master[MF.id]);

    // Calculate grand total including all charges
    const chargesTotal = charges.reduce((sum, charge) => sum + charge.amount, 0);
    const grandTotal = rows.reduce((s, r) => s + (Number(r.total) || 0), 0) || (subtotal + vatTotal + chargesTotal);
    // ✅ PASS NEW PARAMETERS
    const header = await buildHeader(master, brandsText, "Commercial Invoice", orderNumbers, notifyParties, containers, lang, [], customerModels);
    const footer = await buildFooter(master);
    const currency = master[MF.currency] || "SAR";
    const combinedTable = buildCombinedTable_CI(
      rows,
      {
        totalPackages,
        subtotal,
        charges,
        vatTotal,
        grandTotal
      },
      lang,
      currency,
      showHsCode
    );
    const { Document, Paragraph, TextRun, AlignmentType } = w.docx;

    const termsParagraphs = [];
    if (Array.isArray(termsList) && termsList.length) {
      termsList.forEach((t, i) => {
        termsParagraphs.push(new Paragraph({
          bidirectional: rtl,
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          children: [new TextRun({ text: `- ${t}` })]  // ✅ Changed to dash format
        }));
      });
    }

    const additionalTable = buildAdditionalDetailsTable(additionalDetails, "CI", lang);

    const children = [];

    // ✅ ADD 2 BLANK LINES AT THE START
    children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));

    // ✅ 1. ADDITIONAL COMMENTS FIRST (if they exist)
    if (Array.isArray(termsList) && termsList.length) {
      children.push(new Paragraph({
        bidirectional: rtl,
        alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children: [new TextRun({ text: L.additionalComments, bold: true })]
      }));
      children.push(...termsParagraphs);
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
    }

    // ✅ 2. ITEMS TABLE
    // ✅ 2. COMBINED ITEMS & TOTALS TABLE
    children.push(combinedTable);

    // ✅ 4. ADDITIONAL DETAILS TABLE (if exists)
    if (additionalTable) {
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
      children.push(additionalTable);
    }

    return new Document({
      styles: { default: { document: { run: { font: "Arial", size: 18 } } } },
      sections: [{
        headers: { default: header },
        footers: { default: footer },
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: MARGIN_L, right: MARGIN_R }
          },
          bidi: rtl
        },
        children
      }]
    });
  }

  /* =========================
   7) PL BUILDERS
   ========================= */
  function buildConsolidatedTable_PL(containerGroups, plans, lang = "EN", showHsCode = true) {
    const {
      Table, TableRow, TableCell, Paragraph, TextRun,
      WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign
    } = w.docx;
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    const OUTER = { style: BorderStyle.SINGLE, size: 10, color: "000000" };
    const GRID = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
    const TH_FILL = "F7F1C6";
    const HDR_FILL = "D1D5DB";
    const TOTAL_FILL = "B8D4E8";  // Light blue for total row

    // Column widths - removed order column, added optional hsCode
    const CW = showHsCode ? {
      sn: 600, code: 1400, hsCode: 1400, desc: 3200,
      pack: 1300, uom: 900, qty: 1000, liters: 1400,
      net: 1400, gross: 1400
    } : {
      sn: 700, code: 1600, desc: 3600,
      pack: 1400, uom: 1000, qty: 1100, liters: 1500,
      net: 1500, gross: 1600
    };

    const sum = Object.values(CW).reduce((a, b) => a + b, 0);
    const scale = CONTENT_W / sum;
    Object.keys(CW).forEach(k => CW[k] = Math.floor(CW[k] * scale));

    const th = (label, w) => new TableCell({
      width: { size: w, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, color: "auto", fill: TH_FILL },
      borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
      children: [
        new Paragraph({
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.CENTER,
          bidirectional: rtl,
          children: [new TextRun({ text: label, bold: true, font: "Arial", size: 18 })]
        })
      ]
    });

    const td = (val, w, align = "C") => new TableCell({
      width: { size: w, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
      children: [
        new Paragraph({
          alignment:
            align === "L"
              ? (rtl ? AlignmentType.RIGHT : AlignmentType.LEFT)
              : align === "R"
                ? AlignmentType.RIGHT
                : (rtl ? AlignmentType.RIGHT : AlignmentType.CENTER),
          bidirectional: rtl,
          children: [new TextRun({ text: String(val ?? ""), font: "Arial", size: 18 })]
        })
      ]
    });

    // Total row cell with background color
    const tdTotal = (val, w, align = "C") => new TableCell({
      width: { size: w, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, color: "auto", fill: TOTAL_FILL },
      borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
      children: [
        new Paragraph({
          alignment:
            align === "L"
              ? (rtl ? AlignmentType.RIGHT : AlignmentType.LEFT)
              : align === "R"
                ? AlignmentType.RIGHT
                : (rtl ? AlignmentType.RIGHT : AlignmentType.CENTER),
          bidirectional: rtl,
          children: [new TextRun({ text: String(val ?? ""), bold: true, font: "Arial", size: 18 })]
        })
      ]
    });

    // Container total row builder
    const containerTotalRow = (totals) => {
      // Column span depends on whether HS code is shown: sn + code + (hsCode?) + desc + pack + uom
      const colSpan = showHsCode ? 6 : 5;
      const labelWidth = showHsCode
        ? CW.sn + CW.code + CW.hsCode + CW.desc + CW.pack + CW.uom
        : CW.sn + CW.code + CW.desc + CW.pack + CW.uom;
      const totalCells = [
        new TableCell({
          columnSpan: colSpan,
          width: { size: labelWidth, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.CLEAR, color: "auto", fill: TOTAL_FILL },
          borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
          children: [
            new Paragraph({
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.CENTER,
              bidirectional: rtl,
              children: [new TextRun({ text: "Total", bold: true, font: "Arial", size: 18 })]
            })
          ]
        }),
        tdTotal(n(totals.qty, 0), CW.qty, "C"),
        tdTotal(totals.liters > 0 ? n(totals.liters, 2) : "", CW.liters, "C"),
        tdTotal(totals.net > 0 ? n(totals.net, 2) : "", CW.net, "C"),
        tdTotal(totals.gross > 0 ? n(totals.gross, 2) : "", CW.gross, "C")
      ];

      return new TableRow({
        children: rtl ? totalCells.reverse() : totalCells
      });
    };

    // Total columns: sn + code + (hsCode?) + desc + pack + uom + qty + liters + net + gross
    const totalColumns = showHsCode ? 10 : 9;
    const containerHeaderRow = (containerNumber) => new TableRow({
      children: [
        new TableCell({
          columnSpan: totalColumns,
          width: { size: CONTENT_W, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.CLEAR, color: "auto", fill: HDR_FILL },
          borders: { top: GRID, bottom: GRID, left: GRID, right: GRID },
          children: [
            new Paragraph({
              alignment: rtl ? AlignmentType.RIGHT : AlignmentType.CENTER,
              bidirectional: rtl,
              children: [new TextRun({ text: `${L.container}${containerNumber}`, bold: true, font: "Arial", size: 18 })]
            })
          ]
        })
      ]
    });

    // Build header cells - removed Order Number, optional HS Code after Item Code
    const headerCells = showHsCode ? [
      th(L.sn, CW.sn),
      th(L.itemCode, CW.code),
      th(L.hsCode, CW.hsCode),
      th(L.itemDesc, CW.desc),
      th(L.packaging, CW.pack),
      th(L.uom, CW.uom),
      th(L.qty, CW.qty),
      th(L.totalLiters, CW.liters),
      th(L.netWeight, CW.net),
      th(L.grossWeight, CW.gross)
    ] : [
      th(L.sn, CW.sn),
      th(L.itemCode, CW.code),
      th(L.itemDesc, CW.desc),
      th(L.packaging, CW.pack),
      th(L.uom, CW.uom),
      th(L.qty, CW.qty),
      th(L.totalLiters, CW.liters),
      th(L.netWeight, CW.net),
      th(L.grossWeight, CW.gross)
    ];

    const header = new TableRow({
      tableHeader: true,
      children: rtl ? headerCells.reverse() : headerCells
    });

    const rows = [header];
    let serial = 1;

    containerGroups.forEach(group => {
      if (group.containerNumber) {
        rows.push(containerHeaderRow(group.containerNumber));
      }

      // Track totals for this container
      let containerTotals = { qty: 0, liters: 0, net: 0, gross: 0 };

      group.items.forEach(item => {
        const p = plans[item.planId] || {};

        // ✅ Get UOM as number (volume per unit)
        const uomValue = Number(p[LF_PL.uom]) || 0;

        // ✅ Get quantity from container item
        const itemQty = Number(item.qty) || 0;

        // ✅ Get description for COOLANT check
        const description = (p[LF_PL.description] || "").toUpperCase();

        // ✅ Get loaded quantity for packaging weight
        const loadedQty = Number(p[LF_PL.loadedQty]) || 0;

        // ✅ CALCULATION 1: Total Liters = UOM × Qty
        const totalLiters = uomValue * itemQty;

        // ✅ CALCULATION 2: Net Weight (Kgs)
        // If description contains "COOLANT": totalLiters × 1.07
        // Else: totalLiters × 0.9
        let netWeight = 0;
        if (totalLiters > 0) {
          if (description.includes("COOLANT")) {
            netWeight = totalLiters * 1.07;
          } else {
            netWeight = totalLiters * 0.9;
          }
        }

        // ✅ CALCULATION 3: Gross Weight (Kgs)
        // grossWeight = palletsWeight + netWeight + loadedQty (packaging weight)
        const palletsWeight = Number(p[LF_PL.palletWeight]) || 0;
        const packagingWeight = loadedQty; // 1 kg per package assumption
        const grossWeight = palletsWeight + netWeight + packagingWeight;

        // Accumulate totals for container
        containerTotals.qty += itemQty;
        containerTotals.liters += totalLiters;
        containerTotals.net += netWeight;
        containerTotals.gross += grossWeight;

        // Build row cells - removed Order Number, optional HS Code after Item Code
        const rowCells = showHsCode ? [
          td(serial++, CW.sn, "C"),
          td(p[LF_PL.itemCode] || "", CW.code, "C"),
          td(p[LF_PL.hsCode] || "", CW.hsCode, "C"),
          td(p[LF_PL.description] || "", CW.desc, "L"),
          td(p[LF_PL.packaging] || "", CW.pack, "C"),
          td(uomValue > 0 ? n(uomValue, 2) : "", CW.uom, "C"),
          td(n(itemQty, 0), CW.qty, "C"),
          td(totalLiters > 0 ? n(totalLiters, 2) : "", CW.liters, "C"),
          td(netWeight > 0 ? n(netWeight, 2) : "", CW.net, "C"),
          td(grossWeight > 0 ? n(grossWeight, 2) : "", CW.gross, "C")
        ] : [
          td(serial++, CW.sn, "C"),
          td(p[LF_PL.itemCode] || "", CW.code, "C"),
          td(p[LF_PL.description] || "", CW.desc, "L"),
          td(p[LF_PL.packaging] || "", CW.pack, "C"),
          td(uomValue > 0 ? n(uomValue, 2) : "", CW.uom, "C"),
          td(n(itemQty, 0), CW.qty, "C"),
          td(totalLiters > 0 ? n(totalLiters, 2) : "", CW.liters, "C"),
          td(netWeight > 0 ? n(netWeight, 2) : "", CW.net, "C"),
          td(grossWeight > 0 ? n(grossWeight, 2) : "", CW.gross, "C")
        ];

        rows.push(
          new TableRow({
            children: rtl ? rowCells.reverse() : rowCells
          })
        );
      });

      // Add total row for this container
      rows.push(containerTotalRow(containerTotals));
    });

    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: {
        top: OUTER,
        bottom: OUTER,
        left: OUTER,
        right: OUTER,
        insideHorizontal: GRID,
        insideVertical: GRID
      },
      rows
    });
  }

  async function buildDocument_PL(master, brandsText, termsList, containerGroups, plans, additionalDetails, lang = "EN", orderNumbers = [], notifyParties = [], containers = [], showHsCode = true, customerModels = []) {
    await ensureDocx();
    const { Document, Paragraph, TextRun, AlignmentType } = w.docx;
    const L = LABELS[lang.toUpperCase()] || LABELS.EN;
    const rtl = isRTL(lang);

    // Extract unique HS codes from plans
    const hsCodes = [...new Set(
      Object.values(plans)
        .map(p => p[LF_PL.hsCode])
        .filter(Boolean)
    )];

    const header = await buildHeader(master, brandsText, "Packing List", orderNumbers, notifyParties, containers, lang, hsCodes, customerModels);
    const footer = await buildFooter(master);
    const consolidated = buildConsolidatedTable_PL(containerGroups, plans, lang, showHsCode);

    const children = [];

    children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));

    if (Array.isArray(termsList) && termsList.length) {
      children.push(new Paragraph({
        bidirectional: rtl,
        alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children: [new TextRun({ text: L.additionalComments, bold: true })]
      }));
      termsList.forEach((t, i) => {
        children.push(new Paragraph({
          bidirectional: rtl,
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          children: [new TextRun({ text: `- ${t}` })]
        }));
      });
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
    }

    children.push(consolidated);

    const additionalTable = buildAdditionalDetailsTable(additionalDetails, "PL", lang);
    if (additionalTable) {
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
      children.push(new Paragraph({ children: [new TextRun({ text: "\n" })] }));
      children.push(additionalTable);
    }

    return new Document({
      styles: { default: { document: { run: { font: "Arial", size: 18 } } } },
      sections: [{
        headers: { default: header },
        footers: { default: footer },
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: MARGIN_L, right: MARGIN_R }
          },
          bidi: rtl
        },
        children
      }]
    });
  }

  /* =========================
     8) FLOW UPLOAD (shared)
     ========================= */
  async function uploadToFlow({ dclId, blob, language, docType, fileExtension = "docx" }) {
    if (!FLOW_URL) return { ok: false, body: null, status: 0, error: "FLOW_URL is empty" };
    const base64 = await blobToBase64Raw(blob);
    const payload = { id: String(dclId), fileContent: base64, docType, fileExtension, language };
    const approxBytes = base64LenToBytes(base64.length);
    debugLog("POSTing to Flow:", { method: "POST", url: "<hidden>", docType, fileExtension, language, base64Chars: base64.length, approxBytes });

    Loading.start();
    try {
      const r = await fetch(FLOW_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Accept": "application/json" },
        body: JSON.stringify(payload),
        mode: "cors", cache: "no-store", redirect: "follow"
      });
      const raw = await r.text();
      let body = null; try { body = raw ? JSON.parse(raw) : null; } catch { body = raw; }
      debugLog("Flow response:", { status: r.status, ok: r.ok, body });
      if (!r.ok) return { ok: false, status: r.status, body, error: "Flow upload failed" };
      return { ok: true, status: r.status, body };
    } catch (e) {
      debugLog("Flow error:", e);
      return { ok: false, status: 0, body: null, error: String(e && e.message || e) };
    } finally {
      Loading.stop();
    }
  }

  /* =========================
     9) DOCS INDEX & CARD UI
     ========================= */
  const DCL_FIELD = "_cr650_dcl_number_value";
  const CARDS = [
    {
      typeLabel: "Commercial Invoice",
      genSel: "#btnGenCI",
      prevSel: "#btnViewCI",
      regenId: "btnRegenCI",
      radiosName: "ci_lang",
      warningId: "ciWarning"
    },
    {
      typeLabel: "Packing List",
      genSel: "#btnGenPL",
      prevSel: "#btnPrevPL",
      regenId: "btnRegenPL",
      radiosName: "pl_lang",
      warningId: "plWarning"
    }
  ];

  function showDocWarning(warningId) {
    if (!warningId) return;
    const warning = document.getElementById(warningId);
    if (warning) {
      warning.classList.remove('hidden');
    }
  }

  function hideDocWarning(warningId) {
    if (!warningId) return;
    const warning = document.getElementById(warningId);
    if (warning) {
      warning.classList.add('hidden');
    }
  }

  function eqGuid(field, guid) { return field + " eq guid('" + guid + "')"; }
  function valOf(resp) { if (!resp) return []; if (Array.isArray(resp) && resp[0] && typeof resp[0] === "object" && "value" in resp[0]) return resp[0].value || []; if (resp && typeof resp === "object" && "value" in resp) return resp.value || []; return []; }
  function indexByTypeLang(rows) {
    const idx = { "Commercial Invoice": { EN: null, AR: null }, "Packing List": { EN: null, AR: null } };
    const norm = s => (s || "").trim().toUpperCase();
    rows.forEach(r => {
      const t = (r.cr650_doc_type || "").trim();
      if (!idx[t]) return;
      let L = norm(r.cr650_documentlanguage);
      if (L !== "EN" && L !== "AR") L = "EN";
      const cur = idx[t][L];
      const curT = cur ? new Date(cur.modifiedon || cur.createdon || 0).getTime() : 0;
      const newT = new Date(r.modifiedon || r.createdon || 0).getTime();
      if (!cur || newT > curT) idx[t][L] = r;
    });
    return idx;
  }
  function buildDocUrl(filterExpr) {
    const select = [
      "cr650_dcl_documentid", "cr650_doc_type", "cr650_documenturl",
      "cr650_documentlanguage", "_cr650_dcl_number_value", "createdon", "modifiedon"
    ].join(",");
    return DOC_API + "?" + qp({ "$select": select, "$filter": filterExpr, "$orderby": "modifiedon desc", "$top": 100 });
  }
  async function fetchDocs(dclId) {
    const baseTypeFilter = "(cr650_doc_type eq 'Commercial Invoice' or cr650_doc_type eq 'Packing List')";
    // Try bare GUID first (most common working format), then guid() wrapper, then braces
    const filters = [
      `${DCL_FIELD} eq ${dclId} and ${baseTypeFilter}`,
      `${eqGuid(DCL_FIELD, dclId)} and ${baseTypeFilter}`,
      `${DCL_FIELD} eq {${dclId}} and ${baseTypeFilter}`
    ];
    for (const f of filters) {
      try {
        const r = await $.ajax({ url: buildDocUrl(f), headers: { "Accept": "application/json;odata.metadata=none" } });
        if (r && Array.isArray(r.value)) return r;
      } catch { /* try next filter format */ }
    }
    // Fallback: fetch all docs of these types and filter client-side
    try {
      const r = await $.ajax({ url: buildDocUrl(baseTypeFilter), headers: { "Accept": "application/json;odata.metadata=none" } });
      const v = (r && r.value) ? r.value : [];
      return { value: v.filter(row => (row[DCL_FIELD] || "").toLowerCase() === dclId.toLowerCase()) };
    } catch (err) {
      console.error("fetchDocs: all attempts failed", err);
      return { value: [] };
    }
  }
  async function refreshDocs(dclId) {
    if (!dclId || !guidOK(dclId)) return;

    Loading.start();
    try {
      const resp = await fetchDocs(dclId);
      const rows = valOf(resp);
      const idx = indexByTypeLang(rows);
      CARDS.forEach(c => renderCard(c, idx));
    } catch (err) {
      console.error("Failed to refresh docs", err);
    } finally {
      Loading.stop();
    }
  }

  function renderCard(card, docsIndex) {
    const $gen = $(card.genSel);
    const $view = $(card.prevSel);
    if (!$gen.length) return;

    let $regenBtn = $("#" + card.regenId);
    if (!$regenBtn.length) {
      const $regen = $('<button/>', {
        id: card.regenId,
        class: "btn btn-primary hidden",
        title: `Regenerate ${card.typeLabel}`,
        "aria-label": `Regenerate ${card.typeLabel}`,
        type: "button"
      }).append('<i class="fas fa-file-word" aria-hidden="true"></i> Regenerate');
      $gen.after($regen);
      $regenBtn = $("#" + card.regenId);
    }

    const $radios = $(`input[name="${card.radiosName}"]`);
    const getLang = () => {
      const checked = $radios.filter(":checked");
      return (checked.length && checked.val() === "ar") ? "AR" : "EN";
    };

    function setGenerateMode() {
      $gen.removeClass("hidden").attr("data-action", card.typeLabel === "Commercial Invoice" ? "gen-ci" : "gen-pl");
      $view.addClass("hidden").attr("aria-disabled", "true").attr("href", "#");
      if ($regenBtn && $regenBtn.length) {
        $regenBtn.addClass("hidden").off("click");
      }
      hideDocWarning(card.warningId);
    }

    function setViewMode(url) {
      $gen.addClass("hidden").removeAttr("data-action").off("click");
      $view.removeClass("hidden").attr("aria-disabled", "false").attr("href", url);

      $view.off("click").on("click", function (e) {
        if (this.getAttribute("aria-disabled") === "true") {
          e.preventDefault();
          return;
        }
      });

      if ($regenBtn && $regenBtn.length) {
        $regenBtn.removeClass("hidden").off("click").on("click", function (e) {
          e.preventDefault();
          if (card.typeLabel === "Commercial Invoice") {
            generateCI().catch(err => {
              console.error(err);
              alert("Could not regenerate Commercial Invoice.\n" + (err?.message || String(err)));
            });
          } else {
            generatePackingList().catch(err => {
              console.error(err);
              alert("Could not regenerate Packing List.\n" + (err?.message || String(err)));
            });
          }
        });
      }

      showDocWarning(card.warningId);
    }

    function applyFor(lang) {
      const byType = (docsIndex[card.typeLabel] || {});
      const doc = byType[lang];
      const url = doc && doc.cr650_documenturl ? doc.cr650_documenturl : null;
      if (url) {
        setViewMode(url);
      } else {
        setGenerateMode();
      }
    }

    applyFor(getLang());
    $radios.off("change.dcl").on("change.dcl", () => applyFor(getLang()));

    w.addEventListener("dcl:doc:updated", (ev) => {
      const { type, language, fileUrl } = ev.detail || {};
      if (type === card.typeLabel) {
        console.log(`📝 Updating docsIndex for ${type} (${language}):`, fileUrl);

        if (!docsIndex[type]) {
          docsIndex[type] = { EN: null, AR: null };
        }
        docsIndex[type][language] = {
          cr650_documenturl: fileUrl,
          cr650_documentlanguage: language,
          cr650_doc_type: type,
          createdon: new Date().toISOString(),
          modifiedon: new Date().toISOString()
        };

        const currentLang = getLang();
        if (language === currentLang && fileUrl) {
          setViewMode(fileUrl);
        }
      }
    });
  }

  /* =========================
     10) CI ORCHESTRATION
     ========================= */
  let ciInFlight = false;
  function getSelectedLang_CI() { const sel = d.querySelector('input[name="ci_lang"]:checked'); return sel?.value === "ar" ? "ar" : "en"; }
  async function generateCI() {
    try { const btn = d.getElementById("btnGenCI"); if (btn && btn.getAttribute("type") !== "button") btn.setAttribute("type", "button"); } catch { }
    if (ciInFlight) return;
    ciInFlight = true;

    const dclId = getQueryId();
    if (!dclId) { ciInFlight = false; throw new Error("Missing ?id=<guid>."); }

    Loading.start();
    try {
      console.log("=== CI GENERATION START ===");
      debugLog("Starting CI generation for DCL:", dclId);

      const master = await fetchMaster(dclId);
      console.log("📋 Master record fetched:", master);
      console.log("📋 CI Number field:", master[MF.ciNumber]);
      console.log("📋 Auto Number CI field:", master[MF.autoNumberCI]);
      console.log("📋 Country field:", master[MF.country]);

      const containers = await fetchContainers(dclId);
      const items = [];
      for (const c of containers) items.push(...await fetchItemsByContainer(c[CF.id]));
      const planIds = [...new Set(items.map(x => x[IF.planLkp]).filter(Boolean))];
      const plans = {};
      await Promise.all(planIds.map(async id => { const rec = await fetchPlanCI(id); if (rec) plans[id] = rec; }));

      // Fallback: if no items found via containers, fetch loading plans directly
      if (items.length === 0) {
        console.warn("⚠️ CI: No items found via containers, falling back to direct loading plans query");
        const directPlans = await fetchLoadingPlansForCI(dclId);
        console.log("📋 Direct loading plans fetched:", directPlans.length);
        directPlans.forEach(lp => {
          const id = lp[LF_CI.id];
          plans[id] = lp;
          items.push({ [IF.planLkp]: id, [IF.qty]: lp.cr650_loadedquantity || 0 });
        });
      }
      console.log("📦 CI items count:", items.length, "plans count:", Object.keys(plans).length);

      const brandsArray = await fetchBrands(dclId);
      console.log("🏷️ Brands fetched:", brandsArray);
      const brandsText = brandsArray.join(", ");

      const termsList = await fetchTerms(dclId, "CI");
      const additionalDetails = await fetchAdditionalDetails(dclId);

      // ✅ EXTRACT UNIQUE ORDER NUMBERS
      const orderNumbers = [...new Set(
        items.map(it => {
          const p = plans[it[IF.planLkp]] || {};
          return p[LF_CI.orderNo];
        }).filter(Boolean)
      )];

      // ✅ FETCH NOTIFY PARTIES
      const notifyParties = await fetchNotifyParties(dclId);

      // ✅ FETCH CUSTOMER MODELS
      const customerModels = await fetchCustomerModels(dclId);

      // ✅ GENERATE OR USE EXISTING CI NUMBER
      let ciNumber = master[MF.ciNumber];
      console.log("🔢 Existing CI Number:", ciNumber);

      if (!ciNumber) {
        console.log("⚙️ CI Number is null, checking generation requirements...");
        console.log("   - Brands array length:", brandsArray.length);
        console.log("   - Auto Number IC:", master[MF.autoNumberCI]);
        console.log("   - Country:", master[MF.country]);

        if (master[MF.autoNumberCI]) {
          // ✅ Use first brand if available, otherwise default to "Petromin" (P)
          const firstBrand = brandsArray.length > 0 ? brandsArray[0] : "Petromin";
          const autoNumber = master[MF.autoNumberCI];
          const country = master[MF.country];

          if (brandsArray.length === 0) {
            console.log("⚠️ No brands found, using default: Petromin (P)");
          }

          console.log("✅ All requirements met, generating CI number...");
          console.log("   - First Brand:", firstBrand);
          console.log("   - Auto Number:", autoNumber);
          console.log("   - Country:", country);

          ciNumber = generateCINumber(firstBrand, autoNumber, country);
          console.log("📝 Generated CI Number:", ciNumber);

          if (ciNumber) {
            console.log("💾 Saving CI number to Dataverse...");
            await updateCINumber(dclId, ciNumber);
            master[MF.ciNumber] = ciNumber;
            console.log("✅ CI number saved successfully");
          } else {
            console.error("❌ generateCINumber returned null/undefined");
          }
        } else {
          console.warn("⚠️ Cannot generate CI number - Auto Number IC is missing");
        }
      }


      console.log("📄 Final CI Number being used:", ciNumber || master[MF.dcl]);

      const language = getSelectedLang_CI();
      const langUpper = language.toUpperCase();

      const showHsCode = d.getElementById("ciShowHsCode")?.checked ?? true;
      const doc = await buildDocument_CI(
        master,
        items,
        plans,
        brandsText,
        termsList,
        additionalDetails,
        langUpper,
        orderNumbers,
        notifyParties,
        containers,
        showHsCode,
        customerModels
      );

      const blob = await w.docx.Packer.toBlob(doc);

      const fn = `Commercial_Invoice_${master[MF.dcl] || dclId}.docx`;
      console.log("📄 Filename:", fn);

      const url = URL.createObjectURL(blob);
      const a = d.createElement("a"); a.href = url; a.download = fn; d.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);

      const res = await uploadToFlow({
        dclId,
        blob,
        language,
        docType: "Commercial Invoice",
        fileExtension: "docx"
      });

      if (res.ok && res.body) {
        const flowBody = res.body.body || res.body;

        if (flowBody.success && flowBody.fileUrl) {
          console.log("✅ CI uploaded successfully:", flowBody.fileUrl);

          w.dispatchEvent(new CustomEvent("dcl:doc:updated", {
            detail: {
              type: "Commercial Invoice",
              language: langUpper,
              fileUrl: flowBody.fileUrl
            }
          }));

          console.log("✅ UI updated via dcl:doc:updated event");
        } else {
          console.warn("⚠️ Flow returned success but no fileUrl:", res.body);
        }
      } else {
        console.error("❌ Flow upload failed:", res);
        alert("Upload failed. Status: " + (res.status || "unknown"));
      }

      console.log("=== CI GENERATION END ===");

    } finally {
      ciInFlight = false;
      Loading.stop();
    }
  }
  d.addEventListener("click", (e) => {
    const btn = e.target.closest('#btnGenCI[data-action="gen-ci"]');
    if (!btn) return;
    e.preventDefault(); e.stopPropagation();
    generateCI().catch(err => { console.error(err); alert("Could not generate Commercial Invoice.\n" + (err?.message || String(err))); });
  });

  /* =========================
     11) PL ORCHESTRATION
     ========================= */
  let plInFlight = false;
  function getSelectedLang_PL() { const sel = d.querySelector('input[name="pl_lang"]:checked'); return sel?.value === "ar" ? "ar" : "en"; }
  async function generatePackingList() {
    if (plInFlight) return;
    plInFlight = true;
    try { const btn = d.getElementById("btnGenPL"); if (btn && btn.getAttribute("type") !== "button") btn.setAttribute("type", "button"); } catch { }

    const dclId = getQueryId();
    if (!dclId) { plInFlight = false; throw new Error("Missing ?id=<guid>."); }

    w.dispatchEvent(new CustomEvent("dcl:gen:start", { detail: { type: "Packing List" } }));
    Loading.start();
    try {
      const master = await fetchMaster(dclId);
      const containers = await fetchContainers(dclId);

      // ✅ Fetch all container items
      const allItems = [];
      for (const c of containers) {
        const items = await fetchItemsByContainer(c[CF.id]);
        allItems.push(...items);
      }

      // ✅ Fetch all plans
      const planIds = [...new Set(allItems.map(x => x[IF.planLkp]).filter(Boolean))];
      const plans = {};
      await Promise.all([...planIds].map(async pid => {
        const rec = await fetchPlanPL(pid);
        if (rec) plans[pid] = rec;
      }));

      // Fallback: if no items found via containers, fetch loading plans directly
      if (allItems.length === 0) {
        console.warn("⚠️ PL: No items found via containers, falling back to direct loading plans query");
        const directPlans = await fetchLoadingPlansForPL(dclId);
        console.log("📋 Direct loading plans fetched:", directPlans.length);
        directPlans.forEach(lp => {
          const id = lp[LF_PL.id];
          plans[id] = lp;
          allItems.push({ [IF.planLkp]: id, [IF.qty]: lp[LF_PL.loadedQty] || 0 });
        });
      }
      console.log("📦 PL items count:", allItems.length, "plans count:", Object.keys(plans).length);

      // ✅ Group items by container number from loading plans
      const containerMap = new Map();

      allItems.forEach(item => {
        const plan = plans[item[IF.planLkp]];
        if (plan) {
          const containerNumber = plan[LF_PL.containerNumber] || "";

          if (!containerMap.has(containerNumber)) {
            containerMap.set(containerNumber, []);
          }

          containerMap.get(containerNumber).push({
            planId: item[IF.planLkp],
            qty: item[IF.qty]
          });
        }
      });

      // ✅ Convert map to sorted array of groups
      const containerGroups = Array.from(containerMap.entries())
        .sort((a, b) => a[0].localeCompare(b[0]))  // Sort by container number
        .map(([containerNumber, items]) => ({
          containerNumber,
          items
        }));

      console.log("📦 Container groups:", containerGroups);

      const brandsArray = await fetchBrands(dclId);
      const brandsText = brandsArray.join(", ");
      const termsList = await fetchTerms(dclId, "PL");
      const additionalDetails = await fetchAdditionalDetails(dclId);

      const orderNumbers = [...new Set(
        Object.values(plans).map(p => p[LF_PL.orderNo]).filter(Boolean)
      )];

      const notifyParties = await fetchNotifyParties(dclId);
      const customerModels = await fetchCustomerModels(dclId);

      const lang = getSelectedLang_PL();
      const langUpper = lang.toUpperCase();

      // Get HS Code checkbox value
      const showHsCode = d.getElementById("plShowHsCode")?.checked ?? true;

      await ensureDocx();

      const doc = await buildDocument_PL(
        master,
        brandsText,
        termsList,
        containerGroups,
        plans,
        additionalDetails,
        langUpper,
        orderNumbers,
        notifyParties,
        containers,
        showHsCode,
        customerModels
      );

      const blob = await w.docx.Packer.toBlob(doc);

      const friendly = master?.[MF.dcl] || dclId;
      const filename = `Packing_List_${friendly}.docx`;
      const url = URL.createObjectURL(blob);
      const a = d.createElement("a"); a.href = url; a.download = filename; d.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);

      const res = await uploadToFlow({
        dclId,
        blob,
        language: lang,
        docType: "Packing List",
        fileExtension: "docx"
      });

      if (res && res.ok && res.body) {
        const flowBody = res.body.body || res.body;

        if (flowBody.success && flowBody.fileUrl) {
          console.log("✅ PL uploaded successfully:", flowBody.fileUrl);

          w.dispatchEvent(new CustomEvent("dcl:doc:updated", {
            detail: {
              type: "Packing List",
              language: langUpper,
              fileUrl: flowBody.fileUrl
            }
          }));

          console.log("✅ UI updated via dcl:doc:updated event");
        } else {
          console.warn("⚠️ Flow returned success but no fileUrl:", res.body);
        }
      } else if (res && !res.ok) {
        console.error("❌ Flow upload failed:", res);
        alert("Flow upload failed for Packing List.\nStatus: " + res.status + "\n" + (res.error || JSON.stringify(res.body)));
      }

      w.dispatchEvent(new CustomEvent("dcl:gen:done", { detail: { type: "Packing List" } }));
    } catch (err) {
      console.error(err);
      w.dispatchEvent(new CustomEvent("dcl:gen:error", { detail: { type: "Packing List", error: String(err && err.message || err) } }));
      alert("Could not generate Packing List.\n" + (err?.message || String(err)));
    } finally {
      plInFlight = false;
      Loading.stop();
    }
  }
  d.addEventListener("click", (e) => {
    const plBtn = e.target.closest('#btnGenPL[data-action="gen-pl"]');
    if (!plBtn) return;
    e.preventDefault(); e.stopPropagation();
    generatePackingList().catch(err => { console.error(err); alert("Could not generate Packing List.\n" + (err?.message || String(err))); });
  });

  /* =========================
     12) PAGE INIT (cards + nav)
     ========================= */
  function setDclNumberText(text) { const el = d.getElementById("dclNumber"); if (el) el.textContent = text || "-"; }
  const MASTER_API = "/_api/cr650_dcl_masters";
  const MASTER_ID = "cr650_dcl_masterid";
  const MASTER_NUMBER = "cr650_dclnumber";
  const MASTER_STATUS = "cr650_status";  // ✅ ADD THIS

  function buildMasterUrl(filterExpr) {
    const select = [MASTER_ID, MASTER_NUMBER, MASTER_STATUS, "cr650_currencycode"].join(",");
    return MASTER_API + "?" + qp({ "$select": select, "$filter": filterExpr, "$top": 1 });
  }
  async function fetchDclMaster(dclId) {
    const select = [MASTER_ID, MASTER_NUMBER, MASTER_STATUS, "cr650_currencycode"].join(",");
    // Try bare GUID first (most common working format), then guid() wrapper, then braces
    const filters = [
      `${MASTER_ID} eq ${dclId}`,
      `${eqGuid(MASTER_ID, dclId)}`,
      `${MASTER_ID} eq {${dclId}}`
    ];
    for (const f of filters) {
      try {
        const r = await $.ajax({ url: buildMasterUrl(f), headers: { "Accept": "application/json;odata.metadata=none" } });
        if (r?.value?.length) return r;
      } catch { /* try next filter format */ }
    }
    // Fallback: fetch all (with full select) and find by ID
    try {
      const r = await $.ajax({ url: MASTER_API + "?" + qp({ "$select": select, "$top": 50 }), headers: { "Accept": "application/json;odata.metadata=none" } });
      const v = (r && r.value) ? r.value : [];
      const rec = v.find(row => (row[MASTER_ID] || "").toLowerCase() === dclId.toLowerCase());
      return rec ? { value: [rec] } : { value: [] };
    } catch (err) {
      console.error("fetchDclMaster: all attempts failed", err);
      return { value: [] };
    }
  }

  $(function () {
    const dclId = getQueryId();

    if (guidOK(dclId)) {
      rewriteWizardLinksWithId(dclId);
    }
    wireNavLoaderClicks();

    if (!dclId || !guidOK(dclId)) {
      setDclNumberText("-");
      CARDS.forEach(card => hideDocWarning(card.warningId));
      return;
    }

    Loading.start();
    const pMaster = fetchDclMaster(dclId)
      .then(function (resp) {
        const rec = (resp && resp.value && resp.value[0]) ? resp.value[0] : null;
        const number = rec ? (rec[MASTER_NUMBER] || "-") : "-";
        setDclNumberText(number);

        if (rec && rec[MF.status]) {
          DCL_STATUS = rec[MF.status];
          console.log("📋 DCL Status captured from fetchDclMaster:", DCL_STATUS);
        }

        if (rec && rec.cr650_currencycode) {
          CURRENCY_CODE = rec.cr650_currencycode;
          console.log("💱 Currency set from master fetch:", CURRENCY_CODE);
        }
      })
      .catch(function (err) {
        console.error("fetchDclMaster failed:", err);
        setDclNumberText("-");
      })
      .finally(function () {
        Loading.stop();
        lockFormIfSubmitted();
      });

    // Start modules that don't need currency immediately (parallel)
    refreshDocs(dclId);
    initAdditionalDetails().catch(err => console.error("initAdditionalDetails failed:", err));
    initDiscountCharges().catch(err => console.error("initDiscountCharges failed:", err));
    initTermsConditions().catch(err => console.error("initTermsConditions failed:", err));

    // initOrderItems needs CURRENCY_CODE — wait for master fetch first
    pMaster.finally(function () {
      initOrderItems().catch(err => console.error("initOrderItems failed:", err));
    });
  });

})(window, document, window.jQuery);