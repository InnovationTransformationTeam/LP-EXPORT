
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
        }
    };

    const cache = {
        shippinglines: [],
        hscodes: [],
        docmasters: []
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
        const rows = cache[key] || [];
        const $tbody = $(cfg.tableId + " tbody");
        const tpl = _.template($(cfg.templateId).html());

        if (!rows.length) {
            $tbody.html("");
            $(cfg.emptyId).show();
            return;
        }

        $(cfg.emptyId).hide();
        $tbody.html(rows.map(r => tpl(r)).join(""));
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

        $("#refresh-shippinglines").on("click", () => loadList("shippinglines"));
        $("#refresh-hscodes").on("click", () => loadList("hscodes"));
        $("#refresh-docmasters").on("click", () => loadList("docmasters"));

        loadList("shippinglines");
        loadList("hscodes");
        loadList("docmasters");
    });

})();
