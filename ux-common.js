(function () {
    "use strict";

    const pageName = (location.pathname.split("/").pop() || "index.html").replace(/\.html$/i, "") || "index";
    let statusEl = null;
    let statusTimer = null;

    function install() {
        document.body.classList.add("ux-enhanced", `ux-page-${pageName}`);
        applyDarkModePreference();
        ensureSkipLink();
        ensureStatus();
        enhanceBackButtons();
        patchMessageFunctions();
        observeMessageAreas();
        enhanceTables();
        setupKeyboardShortcuts();
        setupGestorNotasHelpers();
        setupActividadModes();
        setupScrollToTop();
        setupDarkModeToggle();
    }

    function ensureSkipLink() {
        if (document.querySelector(".ux-skip-link")) return;
        const link = document.createElement("a");
        link.className = "ux-skip-link";
        link.href = "#contenido";
        link.textContent = "Saltar al contenido";
        link.addEventListener("click", (event) => {
            event.preventDefault();
            const target = document.querySelector("main, .content, .container");
            if (!target) return;
            if (!target.id) target.id = "contenido";
            target.setAttribute("tabindex", "-1");
            target.focus({ preventScroll: true });
            target.scrollIntoView({ block: "start", behavior: "smooth" });
        });
        document.body.prepend(link);
    }

    function ensureStatus() {
        statusEl = document.getElementById("appUxStatus");
        if (statusEl) return;
        statusEl = document.createElement("div");
        statusEl.id = "appUxStatus";
        statusEl.className = "app-ux-status";
        statusEl.setAttribute("role", "status");
        statusEl.setAttribute("aria-live", "polite");
        document.body.appendChild(statusEl);
    }

    function showStatus(text, type = "info", duration = 3200) {
        if (!statusEl || !text) return;
        const cleanText = toActionableMessage(String(text).replace(/\s+/g, " ").trim());
        if (!cleanText) return;

        statusEl.textContent = cleanText;
        statusEl.className = `app-ux-status ${type || "info"} show`;

        if (statusTimer) clearTimeout(statusTimer);
        if (duration > 0) {
            statusTimer = setTimeout(() => {
                statusEl.classList.remove("show");
            }, duration);
        }
    }

    function confirmDialog(message, options = {}) {
        return new Promise((resolve) => {
            const backdrop = document.createElement("div");
            backdrop.className = "ux-modal-backdrop";
            backdrop.innerHTML = `
                <section class="ux-modal" role="dialog" aria-modal="true" aria-labelledby="uxModalTitle">
                    <div class="ux-modal-header">
                        <h2 class="ux-modal-title" id="uxModalTitle">${escapeHtml(options.title || "Confirmar accion")}</h2>
                    </div>
                    <div class="ux-modal-body">${escapeHtml(message || "")}</div>
                    <div class="ux-modal-actions">
                        <button class="ux-modal-cancel" type="button">${escapeHtml(options.cancelText || "Cancelar")}</button>
                        <button class="ux-modal-confirm" type="button">${escapeHtml(options.confirmText || "Aceptar")}</button>
                    </div>
                </section>
            `;

            const cancelBtn = backdrop.querySelector(".ux-modal-cancel");
            const confirmBtn = backdrop.querySelector(".ux-modal-confirm");

            function close(value) {
                document.removeEventListener("keydown", onKeydown);
                backdrop.remove();
                resolve(value);
            }

            function onKeydown(event) {
                if (event.key === "Escape") close(false);
                if (event.key === "Enter" && document.activeElement === confirmBtn) close(true);
            }

            cancelBtn.addEventListener("click", () => close(false));
            confirmBtn.addEventListener("click", () => close(true));
            backdrop.addEventListener("click", (event) => {
                if (event.target === backdrop) close(false);
            });
            document.addEventListener("keydown", onKeydown);
            document.body.appendChild(backdrop);
            cancelBtn.focus();
        });
    }

    function enhanceBackButtons() {
        document.querySelectorAll("button[onclick]").forEach((button) => {
            const handler = button.getAttribute("onclick") || "";
            if (!/history\.back|window\.history\.back/i.test(handler)) return;
            button.removeAttribute("onclick");
            button.title = button.title || "Volver a la pantalla anterior";
            button.addEventListener("click", () => {
                if (window.history.length > 1) {
                    window.history.back();
                } else {
                    window.location.href = "index.html";
                }
            });
        });

        const nav = document.querySelector(".page-nav");
        if (nav) nav.classList.add("ux-sticky-nav");
    }

    function patchMessageFunctions() {
        ["showMessage", "mostrarMensaje"].forEach((name) => {
            const original = window[name];
            if (typeof original !== "function" || original.__uxWrapped) return;
            const wrapped = function (text, type) {
                const result = original.apply(this, arguments);
                showStatus(text, type || "info");
                return result;
            };
            wrapped.__uxWrapped = true;
            window[name] = wrapped;
        });
    }

    function observeMessageAreas() {
        ["message", "saveStatus"].forEach((id) => {
            const element = document.getElementById(id);
            if (!element) return;
            if (!element.getAttribute("role")) element.setAttribute("role", "status");
            if (!element.getAttribute("aria-live")) element.setAttribute("aria-live", "polite");

            const observer = new MutationObserver(() => {
                const text = element.textContent.trim();
                if (!text || element.style.display === "none") return;
                const type = Array.from(element.classList).find((item) =>
                    ["success", "error", "info", "warning", "saving", "saved"].includes(item)
                ) || "info";
                showStatus(text, type);
            });
            observer.observe(element, { childList: true, characterData: true, subtree: true, attributes: true });
        });
    }

    function enhanceTables() {
        document.querySelectorAll(".table-container, .table-wrap, .ra-panel-table-wrap, .resumen-wrap").forEach((wrap) => {
            if (wrap.dataset.uxHint === "1") return;
            wrap.dataset.uxHint = "1";
            requestAnimationFrame(() => {
                if (wrap.scrollWidth <= wrap.clientWidth + 8) return;
                const hint = document.createElement("div");
                hint.className = "ux-table-hint";
                hint.textContent = "Desplaza horizontalmente para ver todas las columnas.";
                wrap.insertAdjacentElement("afterend", hint);
            });
        });
    }

    function setupKeyboardShortcuts() {
        document.addEventListener("keydown", (event) => {
            if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
                const saveButton = document.getElementById("saveButton") ||
                    document.querySelector([
                        "button[onclick*='guardarCambiosForzado']",
                        "button[onclick*='guardarExcel']",
                        "button[onclick*='guardarExcelLocal']",
                        "button[onclick*='descargarExcel']",
                        "button[onclick*='guardarCambios']",
                        "#btnSave",
                    ].join(", "));
                if (saveButton && !saveButton.disabled) {
                    event.preventDefault();
                    saveButton.click();
                    showStatus("Guardando cambios...", "saving", 1800);
                }
            }
        });
    }

    const DARK_KEY = "appDarkMode";

    function applyDarkModePreference() {
        if (localStorage.getItem(DARK_KEY) === "1") {
            document.documentElement.classList.add("dark-mode");
        }
    }

    function setupScrollToTop() {
        const btn = document.createElement("button");
        btn.className = "ux-scroll-top";
        btn.type = "button";
        btn.title = "Volver arriba";
        btn.setAttribute("aria-label", "Volver al inicio de la página");
        btn.textContent = "↑";
        document.body.appendChild(btn);

        btn.addEventListener("click", () => {
            window.scrollTo({ top: 0, behavior: "smooth" });
            document.querySelectorAll(".table-container").forEach(c => c.scrollTo({ top: 0, behavior: "smooth" }));
        });

        function check() {
            btn.classList.toggle("visible", window.scrollY > 300);
        }
        window.addEventListener("scroll", check, { passive: true });
        check();
    }

    function setupDarkModeToggle() {
        const isDark = document.documentElement.classList.contains("dark-mode");
        const btn = document.createElement("button");
        btn.className = "ux-dark-toggle";
        btn.type = "button";
        btn.title = "Alternar modo oscuro";
        btn.textContent = isDark ? "☀️" : "🌙";
        document.body.appendChild(btn);

        btn.addEventListener("click", () => {
            const dark = document.documentElement.classList.toggle("dark-mode");
            localStorage.setItem(DARK_KEY, dark ? "1" : "0");
            btn.textContent = dark ? "☀️" : "🌙";
        });
    }

    function setupGestorNotasHelpers() {
        if (pageName !== "gestor-notas") return;

        const labels = document.querySelector(".progress-bar-labels");
        if (labels && !document.getElementById("jumpFirstErrorButton")) {
            const btn = document.createElement("button");
            btn.id = "jumpFirstErrorButton";
            btn.className = "ux-error-jump";
            btn.type = "button";
            btn.textContent = "Ir al primer error";
            btn.addEventListener("click", focusFirstInvalid);
            labels.appendChild(btn);
        }

        document.addEventListener("keydown", (event) => {
            const target = event.target;
            if (!target.matches || !target.matches(".nota-input, .ce-nota-input")) return;
            const inputs = Array.from(document.querySelectorAll(".nota-input, .ce-nota-input"))
                .filter((input) => !input.disabled && input.offsetParent !== null);
            const idx = inputs.indexOf(target);
            if (idx === -1) return;

            if (event.key === "Enter" || event.key === "ArrowDown") {
                event.preventDefault();
                const next = inputs[idx + 1];
                if (next) focusInput(next);
            } else if (event.key === "ArrowUp") {
                event.preventDefault();
                const prev = inputs[idx - 1];
                if (prev) focusInput(prev);
            }
        });

        const observer = new MutationObserver(updateErrorState);
        observer.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ["class"] });
        document.addEventListener("input", updateErrorState);
        updateErrorState();
    }

    function focusFirstInvalid() {
        const firstInvalid = document.querySelector(".nota-input.invalid, .ce-nota-input.invalid, tr.invalid-note .nota-input");
        if (!firstInvalid) {
            showStatus("No hay errores marcados.", "success");
            return;
        }
        firstInvalid.scrollIntoView({ behavior: "smooth", block: "center" });
        focusInput(firstInvalid);
    }

    function focusInput(input) {
        input.focus();
        if (typeof input.select === "function") input.select();
    }

    function updateErrorState() {
        const hasErrors = Boolean(document.querySelector(".nota-input.invalid, .ce-nota-input.invalid, tr.invalid-note"));
        document.body.classList.toggle("ux-has-errors", hasErrors);
    }

    function setupActividadModes() {
        if (pageName !== "incluir-actividad") return;
        const tabs = Array.from(document.querySelectorAll(".mode-tab"));
        if (!tabs.length) return;

        function setMode(mode) {
            document.body.dataset.activityMode = mode;
            tabs.forEach((tab) => {
                const active = tab.dataset.mode === mode;
                tab.classList.toggle("active", active);
                tab.setAttribute("aria-selected", active ? "true" : "false");
            });
            document.querySelectorAll("[data-mode-panel]").forEach((panel) => {
                panel.hidden = panel.dataset.modePanel !== mode;
            });
        }

        tabs.forEach((tab) => tab.addEventListener("click", () => setMode(tab.dataset.mode)));
        setMode(document.body.dataset.activityMode || "edit");
    }

    function escapeHtml(value) {
        return String(value || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function toActionableMessage(text) {
        const msg = String(text || "").trim();
        if (!msg) return "";
        if (/No se encontr[oó] el Excel|No hay archivo local seleccionado|No se encontr[oó] un Excel activo/i.test(msg)) {
            return `${msg} Abre Inicio y selecciona un archivo .xlsx válido.`;
        }
        if (/permiso|denegado|denied|EACCES/i.test(msg)) {
            return `${msg} Cierra el Excel en otras ventanas y vuelve a intentar guardar.`;
        }
        if (/Error al guardar|Error al abrir|Error al cargar/i.test(msg)) {
            return `${msg} Si persiste, pulsa "Cargar otro archivo" y vuelve a abrir el libro.`;
        }
        return msg;
    }

    window.appUx = {
        showStatus,
        confirm: confirmDialog,
        focusFirstInvalid,
        hintError: toActionableMessage
    };

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", install, { once: true });
    } else {
        install();
    }
})();
