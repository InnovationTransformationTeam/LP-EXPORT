
/* =========================================================
   Admin Lists Control - JS
   ========================================================= */

(function () {
    "use strict";

    if (typeof $ === "undefined" || typeof _ === "undefined") {
        console.error("This page requires jQuery and Underscore.js.");
        return;
    }

    // ------------------------------
    // Processing Bar
    // ------------------------------
    function showProcessing(msg) {
        const $m = $("#processingMsg");
        $m.text(msg || "Processing...");
        $m.show();
    }
    function hideProcessing() {
        $("#processingMsg").hide();
    }

    // ------------------------------
    // Token + safeAjax
    // ------------------------------
    function getToken() {
        return new Promise((resolve, reject) => {
            if (window.shell && typeof window.shell.getTokenDeferred === "function") {
                window.shell.getTokenDeferred().done(resolve).fail(reject);
            } else {
                const tokenEl = document.querySelector('input[name="__RequestVerificationToken"]');
                if (tokenEl && tokenEl.value) resolve(tokenEl.value);
                else resolve(null);
            }
        });
    }

    const webapi = {};
    webapi.safeAjax = async function (ajaxOptions) {
        const token = await getToken();

        ajaxOptions = ajaxOptions || {};
        ajaxOptions.headers = ajaxOptions.headers || {};

        const headers = Object.assign({
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        }, ajaxOptions.headers);

        if (token) headers["__RequestVerificationToken"] = token;

        const method = (ajaxOptions.type || ajaxOptions.method || "GET").toUpperCase();
        let body = ajaxOptions.data || ajaxOptions.body || null;

        if (body && typeof body === "object" && !(body instanceof FormData)) {
            body = JSON.stringify(body);
        }

        const res = await fetch(ajaxOptions.url, {
            method,
            headers,
            credentials: "same-origin",
            body: (method === "GET" || method === "HEAD") ? null : body
        });

        const text = await res.text();
        if (!res.ok) throw new Error(text || res.statusText);

        try { return JSON.parse(text); }
        catch { return text; }
    };

    async function appAjax(opts) {
        if (typeof window.validateLoginSession === "function") {
            try { window.validateLoginSession(); } catch (e) { }
        }
        return webapi.safeAjax(opts);
    }

    // ------------------------------
    // Table Configs
    // ------------------------------
    const LISTS = {
        shippinglines: {
            key: "shippinglines",
            title: "Shipping Lines",
            entitySet: "cr650_dcl_shippinglines",
            idField: "cr650_dcl_shippinglineid",
            fields: ["cr650_shipping_line"],
            select: "cr650_shipping_line,cr650_dcl_shippinglineid",
            templateId: "#tpl-row-shippinglines",
            tableId: "#table-shippinglines",
            emptyId: "#empty-shippinglines",
            mapRow: (r) => ({
                id: r.cr650_dcl_shippinglineid,
                etag: r["@odata.etag"] || "",
                shipping_line: r.cr650_shipping_line || ""
            })
        },
        hscodes: {
            key: "hscodes",
            title: "HS Codes",
            entitySet: "cr650_hscodes",
            idField: "cr650_hscodeid",
            fields: ["cr650_itemnumber", "cr650_itemdescription", "cr650_hscode", "cr650_category"],
            select: "cr650_itemnumber,cr650_itemdescription,cr650_hscode,cr650_category,cr650_hscodeid",
            templateId: "#tpl-row-hscodes",
            tableId: "#table-hscodes",
            emptyId: "#empty-hscodes",
            mapRow: (r) => ({
                id: r.cr650_hscodeid,
                etag: r["@odata.etag"] || "",
                cr650_itemnumber: r.cr650_itemnumber || "",
                cr650_itemdescription: r.cr650_itemdescription || "",
                cr650_hscode: r.cr650_hscode || "",
                cr650_category: r.cr650_category || ""
            })
        },
        docmasters: {
            key: "docmasters",
            title: "Document Masters",
            entitySet: "cr650_documents_masters",
            idField: "cr650_documents_masterid",
            fields: ["cr650_documentname"],
            select: "cr650_documentname,cr650_documents_masterid",
            templateId: "#tpl-row-docmasters",
            tableId: "#table-docmasters",
            emptyId: "#empty-docmasters",
            mapRow: (r) => ({
                id: r.cr650_documents_masterid,
                etag: r["@odata.etag"] || "",
                cr650_documentname: r.cr650_documentname || ""
            })
        },
        currencies: {
            key: "currencies",
            title: "Currencies",
            entitySet: "cr650_dclcurrencieses",
            idField: "cr650_dclcurrenciesid",
            fields: ["cr650_currencycode", "cr650_currencyname", "cr650_sortorder"],
            select: "cr650_currencycode,cr650_currencyname,cr650_sortorder,cr650_dclcurrenciesid",
            templateId: "#tpl-row-currencies",
            tableId: "#table-currencies",
            emptyId: "#empty-currencies",
            mapRow: (r) => ({
                id: r.cr650_dclcurrenciesid,
                etag: r["@odata.etag"] || "",
                cr650_currencycode: r.cr650_currencycode || "",
                cr650_currencyname: r.cr650_currencyname || "",
                cr650_sortorder: (r.cr650_sortorder == null) ? null : Number(r.cr650_sortorder)
            }),
            // Sort null sortorders to the bottom, then by code
            sortRows: (rows) => rows.slice().sort((a, b) => {
                const aOrd = (a.cr650_sortorder == null) ? Number.MAX_SAFE_INTEGER : a.cr650_sortorder;
                const bOrd = (b.cr650_sortorder == null) ? Number.MAX_SAFE_INTEGER : b.cr650_sortorder;
                if (aOrd !== bOrd) return aOrd - bOrd;
                return (a.cr650_currencycode || "").localeCompare(b.cr650_currencycode || "");
            })
        }
    };

    const cache = {
        shippinglines: [],
        hscodes: [],
        docmasters: [],
        currencies: []
    };

    // ------------------------------
    // Load + Render
    // ------------------------------
    async function loadList(key) {
        const cfg = LISTS[key];
        if (!cfg) return;

        showProcessing("Loading " + cfg.title + "...");
        try {
            const url = "/_api/" + cfg.entitySet + "?$select=" + encodeURIComponent(cfg.select);
            const res = await appAjax({ url, method: "GET" });

            const rows = (res && res.value) ? res.value.map(cfg.mapRow) : [];
            cache[key] = rows;

            renderList(key);
        } catch (e) {
            console.error("Load error", key, e);
            alert("Failed to load " + cfg.title + ".\n\n" + (e.message || e));
        } finally {
            hideProcessing();
        }
    }

    function renderList(key) {
        const cfg = LISTS[key];
        let rows = cache[key] || [];
        const $tbody = $(cfg.tableId + " tbody");
        const tpl = _.template($(cfg.templateId).html());

        if (!rows.length) {
            $tbody.html("");
            $(cfg.emptyId).show();
            return;
        }

        // Apply optional per-list sort (used by currencies)
        if (typeof cfg.sortRows === "function") {
            rows = cfg.sortRows(rows);
            cache[key] = rows; // persist sorted order so move-up/down sees correct neighbors
        }

        $(cfg.emptyId).hide();
        $tbody.html(rows.map((r, idx) => tpl(Object.assign({}, r, {
            isFirst: idx === 0,
            isLast: idx === rows.length - 1
        }))).join(""));
    }

    // ------------------------------
    // Add
    // ------------------------------
    async function addShippingLine() {
        const name = ($("#add-shippinglines-name").val() || "").trim();
        if (!name) {
            alert("Please enter a shipping line name.");
            return;
        }

        await createRecord("shippinglines", { cr650_shipping_line: name });
        $("#add-shippinglines-name").val("");
    }

    async function addHSCode() {
        const payload = {
            cr650_itemnumber: ($("#add-hscodes-itemnumber").val() || "").trim(),
            cr650_itemdescription: ($("#add-hscodes-itemdescription").val() || "").trim(),
            cr650_hscode: ($("#add-hscodes-hscode").val() || "").trim(),
            cr650_category: ($("#add-hscodes-category").val() || "").trim()
        };

        await createRecord("hscodes", payload);

        $("#add-hscodes-itemnumber").val("");
        $("#add-hscodes-itemdescription").val("");
        $("#add-hscodes-hscode").val("");
        $("#add-hscodes-category").val("");
    }

    async function addDocMaster() {
        const name = ($("#add-docmasters-name").val() || "").trim();
        if (!name) {
            alert("Please enter a document name.");
            return;
        }

        await createRecord("docmasters", { cr650_documentname: name });
        $("#add-docmasters-name").val("");
    }

    async function addCurrency() {
        const code = ($("#add-currencies-code").val() || "").trim().toUpperCase();
        const name = ($("#add-currencies-name").val() || "").trim();

        if (!code) {
            alert("Please enter a currency code (e.g., USD).");
            return;
        }
        if (code.length !== 3 || !/^[A-Z]{3}$/.test(code)) {
            alert("Currency code must be exactly 3 letters (ISO 4217 format), e.g., USD, SAR.");
            return;
        }

        // Duplicate-code check before POST (Dataverse returns 412 on alternate-key violation)
        const existing = (cache.currencies || []).find(r => (r.cr650_currencycode || "").toUpperCase() === code);
        if (existing) {
            alert("A currency with code " + code + " already exists.");
            return;
        }

        // New row goes to the bottom of the order (highest sortorder + 1)
        const currentMax = (cache.currencies || []).reduce((m, r) => {
            const n = (r.cr650_sortorder == null) ? 0 : r.cr650_sortorder;
            return n > m ? n : m;
        }, 0);

        const payload = {
            cr650_currencycode: code,
            cr650_currencyname: name || code,
            cr650_sortorder: currentMax + 1,
            // cr650_name is Dataverse's required primary column — mirror the code for a sensible label
            cr650_name: code
        };

        await createRecord("currencies", payload);

        $("#add-currencies-code").val("");
        $("#add-currencies-name").val("");
    }

    async function createRecord(key, data) {
        const cfg = LISTS[key];
        showProcessing("Adding record...");
        try {
            const url = "/_api/" + cfg.entitySet;
            await appAjax({
                url,
                method: "POST",
                data
            });
            await loadList(key);
        } catch (e) {
            console.error("Create error", key, e);
            alert("Failed to add record.\n\n" + (e.message || e));
        } finally {
            hideProcessing();
        }
    }

    // ------------------------------
    // Inline Edit Helpers
    // ------------------------------
    function enterEditMode($tr) {
        $tr.addClass("editing");
        $tr.find(".cell-input").prop("disabled", false);

        const orig = {};
        $tr.find(".cell-input").each(function () {
            const $inp = $(this);
            const field = $inp.data("field") || "__single__";
            orig[field] = $inp.val();
        });
        $tr.data("orig", orig);

        $tr.find(".view-mode").addClass("hidden");
        $tr.find(".edit-mode").removeClass("hidden");
    }

    function exitEditMode($tr, restore) {
        if (restore) {
            const orig = $tr.data("orig") || {};
            $tr.find(".cell-input").each(function () {
                const $inp = $(this);
                const field = $inp.data("field") || "__single__";
                if (orig.hasOwnProperty(field)) $inp.val(orig[field]);
            });
        }

        $tr.removeClass("editing");
        $tr.find(".cell-input").prop("disabled", true);
        $tr.find(".edit-mode").addClass("hidden");
        $tr.find(".view-mode").removeClass("hidden");
        $tr.removeData("orig");
    }

    // ------------------------------
    // Save Update
    // ------------------------------
    async function saveRow(key, $tr) {
        const cfg = LISTS[key];
        const id = $tr.attr("data-id");
        const etag = $tr.attr("data-etag") || "*";

        if (!id) {
            alert("Missing record id.");
            return;
        }

        const payload = {};

        if (key === "shippinglines") {
            const val = ($tr.find(".cell-input").val() || "").trim();
            if (!val) {
                alert("Shipping line name cannot be empty.");
                return;
            }
            payload.cr650_shipping_line = val;
        }

        if (key === "docmasters") {
            const val = ($tr.find(".cell-input").val() || "").trim();
            if (!val) {
                alert("Document name cannot be empty.");
                return;
            }
            payload.cr650_documentname = val;
        }

        if (key === "hscodes") {
            $tr.find(".cell-input").each(function () {
                const $inp = $(this);
                const field = $inp.data("field");
                if (field) payload[field] = ($inp.val() || "").trim();
            });
        }

        if (key === "currencies") {
            $tr.find(".cell-input").each(function () {
                const $inp = $(this);
                const field = $inp.data("field");
                if (!field) return;
                let val = ($inp.val() || "").trim();
                if (field === "cr650_currencycode") val = val.toUpperCase();
                payload[field] = val;
            });
            if (!payload.cr650_currencycode || !/^[A-Z]{3}$/.test(payload.cr650_currencycode)) {
                alert("Currency code must be exactly 3 letters (ISO 4217), e.g., USD.");
                return;
            }
            // Keep cr650_name mirrored to the code so lookups/references display something meaningful
            payload.cr650_name = payload.cr650_currencycode;
            if (!payload.cr650_currencyname) payload.cr650_currencyname = payload.cr650_currencycode;
        }

        showProcessing("Saving changes...");
        try {
            const url = "/_api/" + cfg.entitySet + "(" + id + ")";
            await appAjax({
                url,
                method: "PATCH",
                headers: { "If-Match": etag || "*" },
                data: payload
            });

            await loadList(key);
        } catch (e) {
            console.error("Update error", key, e);
            alert("Failed to update record.\n\n" + (e.message || e));
        } finally {
            hideProcessing();
        }
    }

    // ------------------------------
    // Move Up / Down (currencies only)
    // Swaps cr650_sortorder with the neighbor row in the currently-sorted list.
    // ------------------------------
    async function moveCurrency($tr, direction) {
        const cfg = LISTS.currencies;
        const id = $tr.attr("data-id");
        if (!id) return;

        const rows = cache.currencies || [];
        const idx = rows.findIndex(r => r.id === id);
        if (idx < 0) return;

        const neighborIdx = direction === "up" ? idx - 1 : idx + 1;
        if (neighborIdx < 0 || neighborIdx >= rows.length) return;

        const cur = rows[idx];
        const neighbor = rows[neighborIdx];

        // Pick unambiguous sortorder values. If either is null, assign sequential numbers
        // for the whole list so moves always work predictably.
        let curOrder = cur.cr650_sortorder;
        let neighborOrder = neighbor.cr650_sortorder;

        if (curOrder == null || neighborOrder == null || curOrder === neighborOrder) {
            // Renumber everyone from 1..N in current visual order, then swap
            showProcessing("Normalizing order...");
            try {
                for (let i = 0; i < rows.length; i++) {
                    if (rows[i].cr650_sortorder !== i + 1) {
                        await appAjax({
                            url: "/_api/" + cfg.entitySet + "(" + rows[i].id + ")",
                            method: "PATCH",
                            headers: { "If-Match": "*" },
                            data: { cr650_sortorder: i + 1 }
                        });
                        rows[i].cr650_sortorder = i + 1;
                    }
                }
            } catch (e) {
                console.error("Normalize error", e);
                alert("Failed to normalize order.\n\n" + (e.message || e));
                hideProcessing();
                await loadList("currencies");
                return;
            }
            curOrder = rows[idx].cr650_sortorder;
            neighborOrder = rows[neighborIdx].cr650_sortorder;
        } else {
            showProcessing(direction === "up" ? "Moving up..." : "Moving down...");
        }

        try {
            // Swap the two sortorder values
            await appAjax({
                url: "/_api/" + cfg.entitySet + "(" + cur.id + ")",
                method: "PATCH",
                headers: { "If-Match": "*" },
                data: { cr650_sortorder: neighborOrder }
            });
            await appAjax({
                url: "/_api/" + cfg.entitySet + "(" + neighbor.id + ")",
                method: "PATCH",
                headers: { "If-Match": "*" },
                data: { cr650_sortorder: curOrder }
            });

            await loadList("currencies");
        } catch (e) {
            console.error("Reorder error", e);
            alert("Failed to change order.\n\n" + (e.message || e));
            await loadList("currencies"); // refresh to a consistent state
        } finally {
            hideProcessing();
        }
    }

    // ------------------------------
    // Delete
    // ------------------------------
    async function deleteRow(key, $tr) {
        const cfg = LISTS[key];
        const id = $tr.attr("data-id");
        const etag = $tr.attr("data-etag") || "*";

        if (!id) {
            alert("Missing record id.");
            return;
        }

        const ok = confirm("Are you sure you want to delete this record?");
        if (!ok) return;

        showProcessing("Deleting record...");
        try {
            const url = "/_api/" + cfg.entitySet + "(" + id + ")";
            await appAjax({
                url,
                method: "DELETE",
                headers: { "If-Match": etag || "*" }
            });
            await loadList(key);
        } catch (e) {
            console.error("Delete error", key, e);
            alert("Failed to delete record.\n\n" + (e.message || e));
        } finally {
            hideProcessing();
        }
    }

    // ------------------------------
    // Event Binding
    // ------------------------------
    function bindRowEvents(key) {
        const cfg = LISTS[key];
        const $table = $(cfg.tableId);

        $table.off("click.admin");

        $table.on("click.admin", ".btn-edit", function () {
            const $tr = $(this).closest("tr");
            enterEditMode($tr);
        });

        $table.on("click.admin", ".btn-cancel", function () {
            const $tr = $(this).closest("tr");
            exitEditMode($tr, true);
        });

        $table.on("click.admin", ".btn-save", async function () {
            const $tr = $(this).closest("tr");
            await saveRow(key, $tr);
        });

        $table.on("click.admin", ".btn-delete", async function () {
            const $tr = $(this).closest("tr");
            await deleteRow(key, $tr);
        });

        if (key === "currencies") {
            $table.on("click.admin", ".btn-move-up", async function () {
                if ($(this).is(":disabled")) return;
                await moveCurrency($(this).closest("tr"), "up");
            });
            $table.on("click.admin", ".btn-move-down", async function () {
                if ($(this).is(":disabled")) return;
                await moveCurrency($(this).closest("tr"), "down");
            });
        }
    }

    // ------------------------------
    // Tabs
    // ------------------------------
    function initTabs() {
        $(".tab-btn").on("click", function () {
            const tab = $(this).data("tab");

            $(".tab-btn").removeClass("active");
            $(this).addClass("active");

            $(".tab-panel").removeClass("active");
            $("#panel-" + tab).addClass("active");
        });
    }

    // ------------------------------
    // Init
    // ------------------------------
    $(function () {
        initTabs();

        Object.keys(LISTS).forEach(key => bindRowEvents(key));

        $("#add-shippinglines-btn").on("click", addShippingLine);
        $("#add-hscodes-btn").on("click", addHSCode);
        $("#add-docmasters-btn").on("click", addDocMaster);
        $("#add-currencies-btn").on("click", addCurrency);

        $("#refresh-shippinglines").on("click", () => loadList("shippinglines"));
        $("#refresh-hscodes").on("click", () => loadList("hscodes"));
        $("#refresh-docmasters").on("click", () => loadList("docmasters"));
        $("#refresh-currencies").on("click", () => loadList("currencies"));

        loadList("shippinglines");
        loadList("hscodes");
        loadList("docmasters");
        loadList("currencies");
    });

})();
