// ==UserScript==
// @name        Global Reader Theme
// @namespace   Violentmonkey Scripts
// @version     3.2.0
//
// @match       *://*/*
// @grant       none
//
// @author      -
// @description Per-site reader theme with element picker
// ==/UserScript==

(function () {
    if (window.__globalReaderThemeInstalled) return;
    window.__globalReaderThemeInstalled = true;

    const ROOT_ID = "global-reader-theme-root";
    const FOCUS_CLASS = "__reader-focus";
    const STORAGE_PREFIX = "globalReaderTheme:";
    const FOCUS_PREFIX = "globalReaderFocus:";

    const DEFAULT_COLORS = {
        bgHue: 35,
        bgLightness: 92,
        textLightness: 15
    };

    const PRESETS = {
        "www.webnovel.com": {
            configured: true,
            colors: { ...DEFAULT_COLORS },
            selectors: {
                page: ["html", "body"],
                panel: [".cha-page", "._color3 .cha-page"],
                text: [
                    "._color3 .cha-content",
                    ".cha-content",
                    ".cha-content p",
                    ".cha-content span"
                ],
                hide: ["header.cha-header", "div.cha-fly"]
            },
            extraCss: t => `
                body { --bc_dark_primary:${t.panel} !important; }
                .cha-page { border:1px solid ${t.border} !important; }
                .cha-page * { border-color:${t.border} !important; }
            `
        }
    };

    const host = location.hostname;
    const storageKey = STORAGE_PREFIX + host;
    const focusKey = FOCUS_PREFIX + host;

    function emptyConfig() {
        return {
            configured: false,
            colors: { ...DEFAULT_COLORS },
            selectors: {
                page: ["html", "body"],
                panel: [],
                text: [],
                hide: []
            },
            extraCss: null
        };
    }

    function loadConfig() {
        const preset = PRESETS[host];
        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || "null");
            if (saved) {
                return {
                    ...emptyConfig(),
                    ...preset,
                    ...saved,
                    colors: {
                        ...DEFAULT_COLORS,
                        ...(preset?.colors || {}),
                        ...(saved.colors || {})
                    },
                    selectors: {
                        ...emptyConfig().selectors,
                        ...(preset?.selectors || {}),
                        ...(saved.selectors || {})
                    },
                    extraCss: saved.extraCss || preset?.extraCss || null
                };
            }
        } catch {
            // ignore corrupt storage
        }
        if (preset) {
            return {
                ...emptyConfig(),
                ...preset,
                colors: { ...DEFAULT_COLORS, ...preset.colors },
                selectors: { ...emptyConfig().selectors, ...preset.selectors }
            };
        }
        return emptyConfig();
    }

    let config = loadConfig();

    function saveConfig() {
        const payload = {
            configured: config.configured,
            colors: config.colors,
            selectors: config.selectors
        };
        localStorage.setItem(storageKey, JSON.stringify(payload));
    }

    function buildAncestorChain(el) {
        const ancestors = [];
        let parent = el.parentElement;

        while (parent && parent !== document.documentElement) {
            const sel = parentSelector(parent);
            if (sel) ancestors.push(sel);
            parent = parent.parentElement;
        }

        return ancestors;
    }

    function normalizeSelectorEntry(entry) {
        if (!entry) return null;
        if (typeof entry === "string") {
            return { base: entry, ancestors: [], depth: 0 };
        }
        return {
            base: entry.base || "",
            ancestors: entry.ancestors || [],
            depth: entry.depth || 0
        };
    }

    function resolveSelector(entry) {
        const { base, ancestors, depth } = normalizeSelectorEntry(entry);
        if (!base) return "";

        if (!depth || !ancestors.length) return base;

        const prefix = ancestors
            .slice(0, depth)
            .reverse()
            .join(" ");

        return prefix ? `${prefix} ${base}` : base;
    }

    function joinSelectors(list) {
        return (list || []).map(resolveSelector).filter(Boolean).join(", ");
    }

    const ROLE_LABELS = {
        page: "Page",
        panel: "Panel",
        text: "Text",
        hide: "Hide"
    };

    function getRoleList(role) {
        return config.selectors[role] || [];
    }

    function appendRule(role, picked) {
        if (!config.selectors[role]) config.selectors[role] = [];
        const resolved = resolveSelector(picked);
        const duplicate = config.selectors[role].some(
            entry => resolveSelector(entry) === resolved
        );
        if (!duplicate) config.selectors[role].push(picked);
    }

    function removeRule(role, index) {
        config.selectors[role].splice(index, 1);
    }

    function setRuleDepth(role, index, delta) {
        const entry = config.selectors[role][index];
        if (!entry || typeof entry === "string") return;

        const max = entry.ancestors?.length || 0;
        entry.depth = clamp((entry.depth || 0) + delta, 0, max);
        applyTheme();
    }

    function markConfiguredIfReady(role) {
        if (
            role === "text" ||
            role === "panel" ||
            (role === "page" && config.selectors.text.length)
        ) {
            config.configured = true;
        }
    }

    function hsl(h, s, l) {
        return `hsl(${h}, ${s}%, ${l}%)`;
    }

    function clamp(v, min, max) {
        return Math.max(min, Math.min(max, v));
    }

    function getTheme() {
        const { bgHue, bgLightness, textLightness } = config.colors;
        return {
            page: hsl(bgHue, 35, bgLightness),
            panel: hsl(bgHue, 30, clamp(bgLightness + 4, 0, 100)),
            border: hsl(bgHue, 25, clamp(bgLightness - 10, 0, 100)),
            text: hsl(bgHue, 20, textLightness),
            muted: hsl(bgHue, 15, clamp(textLightness + 25, 0, 100)),
            accent: hsl(bgHue, 60, 45)
        };
    }

    function escapeCss(value) {
        if (typeof CSS !== "undefined" && CSS.escape) {
            return CSS.escape(value);
        }
        return value.replace(/([^\w-])/g, "\\$1");
    }

    function countMatches(selector) {
        try {
            return document.querySelectorAll(selector).length;
        } catch {
            return 0;
        }
    }

    function isBadClass(name) {
        if (!name) return true;
        if (/^(active|open|hidden|focus|selected|hover)$/i.test(name)) return true;
        if (/^(ng-|css-|jsx-|sc-|emotion-|svelte-|chakra-|mui-|tw-)/i.test(name)) {
            return true;
        }
        if (/^_[a-zA-Z0-9]{5,}$/.test(name)) return true;
        if (/__[a-z0-9]{5,}$/i.test(name)) return true;
        if (name.length >= 16 && /^[a-zA-Z0-9]+$/.test(name)) return true;
        return false;
    }

    function stableClasses(el) {
        return [...el.classList].filter(c => c && !isBadClass(c));
    }

    function parentSelector(el) {
        if (el.id) return `#${escapeCss(el.id)}`;

        const stable = stableClasses(el);
        if (stable.length) {
            return `${el.tagName.toLowerCase()}.${stable.map(escapeCss).join(".")}`;
        }

        return null;
    }

    function generateTextSelector(el) {
        const childTag = el.tagName.toLowerCase();
        let parent = el.parentElement;
        let fallback = null;

        while (parent && parent !== document.documentElement) {
            const base = parentSelector(parent);
            if (base) {
                const sel = `${base} ${childTag}`;
                const matches = countMatches(sel);
                if (matches >= 2) return sel;
                if (matches === 1 && !fallback) fallback = sel;
            }
            parent = parent.parentElement;
        }

        return fallback || childTag;
    }

    function generateSelector(el, role) {
        if (!el || el.nodeType !== 1) return null;
        if (el === document.documentElement) return "html";
        if (el === document.body) return "body";

        if (role === "text") {
            return generateTextSelector(el);
        }

        if (el.id) {
            const byId = `#${escapeCss(el.id)}`;
            if (countMatches(byId) === 1) return byId;
        }

        const tag = el.tagName.toLowerCase();
        const classes = stableClasses(el);

        for (let n = Math.min(3, classes.length); n >= 1; n--) {
            const sel = `${tag}.${classes.slice(0, n).map(escapeCss).join(".")}`;
            const matches = countMatches(sel);
            if (matches >= 1 && matches <= 25) return sel;
        }

        const parts = [];
        let node = el;
        while (node && node.nodeType === 1 && node !== document.documentElement) {
            let part = node.tagName.toLowerCase();

            if (node.id) {
                parts.unshift(`#${escapeCss(node.id)}`);
                break;
            }

            const parent = node.parentElement;
            if (parent) {
                const siblings = [...parent.children].filter(
                    child => child.tagName === node.tagName
                );
                if (siblings.length > 1) {
                    part += `:nth-of-type(${siblings.indexOf(node) + 1})`;
                }
            }

            const nodeStable = stableClasses(node);
            if (nodeStable.length) {
                part += `.${escapeCss(nodeStable[0])}`;
            }

            parts.unshift(part);
            node = parent;
            if (parts.length >= 6) break;
        }

        const path = parts.join(" > ");
        if (countMatches(path) >= 1) return path;

        return tag;
    }

    function pickSelector(target, role) {
        const base = generateSelector(target, role);
        if (!base) return null;

        return {
            base,
            ancestors: buildAncestorChain(target),
            depth: 0
        };
    }

    function buildCss() {
        if (!config.configured) return "";

        const t = getTheme();
        const { page, panel, text, hide } = config.selectors;
        const rules = [];

        const pageSel = joinSelectors(page);
        if (pageSel) {
            rules.push(`${pageSel} { background:${t.page} !important; }`);
        }

        const panelSel = joinSelectors(panel);
        if (panelSel) {
            rules.push(
                `${panelSel} { background:${t.panel} !important; color:${t.text} !important; border-color:${t.border} !important; }`
            );
        }

        const textSel = joinSelectors(text);
        if (textSel) {
            rules.push(`${textSel} { color:${t.text} !important; }`);
        }

        rules.push(`a { color:${t.accent} !important; }`);

        const hideSel = joinSelectors(hide);
        if (hideSel) {
            rules.push(
                `body.${FOCUS_CLASS} { overflow-x:hidden !important; }`,
                hide
                    .map(s => `body.${FOCUS_CLASS} ${resolveSelector(s)}`)
                    .join(", ") + " { display:none !important; }"
            );
        }

        if (typeof config.extraCss === "function") {
            rules.push(config.extraCss(t));
        }

        return rules.join("\n");
    }

    const style = document.createElement("style");
    style.id = "global-reader-theme-style";
    document.documentElement.appendChild(style);

    function applyTheme() {
        style.textContent = buildCss();
        saveConfig();
        refreshUi();
    }

    function setFocusMode(enabled) {
        if (document.body) {
            document.body.classList.toggle(FOCUS_CLASS, enabled);
        }
        localStorage.setItem(focusKey, enabled ? "true" : "false");
        if (focusToggle) {
            focusToggle.textContent = enabled ? "Zen Mode: On" : "Zen Mode: Off";
            focusToggle.style.background = enabled ? "#16a34a" : "#111827";
        }
        style.textContent = buildCss();
    }

    function isOurElement(el) {
        return el && el.closest && el.closest(`#${ROOT_ID}`);
    }

    function startPicker(role, label, options, onDone) {
        if (typeof options === "function") {
            onDone = options;
            options = {};
        }

        const { mode = "append", index = null } = options;
        const hint = document.createElement("div");
        hint.textContent = `Click the ${label}. Esc to cancel.`;
        hint.style.cssText = `
            position:fixed; top:16px; left:50%; transform:translateX(-50%);
            z-index:2147483647; background:#111827; color:#fff;
            padding:10px 16px; border-radius:999px; font:600 14px Inter,Arial,sans-serif;
            box-shadow:0 8px 24px rgba(0,0,0,.3); pointer-events:none;
        `;

        const highlight = document.createElement("div");
        highlight.style.cssText = `
            position:fixed; pointer-events:none; z-index:2147483646;
            outline:2px solid #3b82f6; outline-offset:2px;
            background:rgba(59,130,246,.12); display:none;
        `;

        document.body.appendChild(hint);
        document.body.appendChild(highlight);

        let active = true;

        function cleanup() {
            if (!active) return;
            active = false;
            hint.remove();
            highlight.remove();
            document.removeEventListener("mousemove", onMove, true);
            document.removeEventListener("click", onClick, true);
            document.removeEventListener("keydown", onKey, true);
        }

        function targetAt(x, y) {
            const elements = document.elementsFromPoint(x, y);
            for (const el of elements) {
                if (!isOurElement(el) && el !== hint && el !== highlight) {
                    return el;
                }
            }
            return null;
        }

        function onMove(e) {
            const target = targetAt(e.clientX, e.clientY);
            if (!target) {
                highlight.style.display = "none";
                return;
            }
            const rect = target.getBoundingClientRect();
            highlight.style.display = "block";
            highlight.style.left = `${rect.left}px`;
            highlight.style.top = `${rect.top}px`;
            highlight.style.width = `${rect.width}px`;
            highlight.style.height = `${rect.height}px`;
        }

        function onClick(e) {
            e.preventDefault();
            e.stopPropagation();
            const target = targetAt(e.clientX, e.clientY);
            if (!target) return;
            const picked = pickSelector(target, role);
            if (!picked) return;

            if (mode === "replace" && index !== null) {
                config.selectors[role][index] = picked;
            } else {
                appendRule(role, picked);
            }

            markConfiguredIfReady(role);
            cleanup();
            applyTheme();
            if (onDone) onDone(picked.base);
        }

        function onKey(e) {
            if (e.key === "Escape") {
                e.preventDefault();
                cleanup();
            }
        }

        document.addEventListener("mousemove", onMove, true);
        document.addEventListener("click", onClick, true);
        document.addEventListener("keydown", onKey, true);
    }

    function shortSelector(sel) {
        if (!sel) return "not set";
        return sel.length > 42 ? `${sel.slice(0, 39)}...` : sel;
    }

    const depthBtnStyle = `
        border:none; background:#e5e7eb; border-radius:4px;
        width:22px; height:22px; cursor:pointer; font-size:14px; line-height:1;
    `;

    function renderRuleRow(role, entry, index) {
        const resolved = resolveSelector(entry);
        const isObject = entry && typeof entry === "object";
        const depth = isObject ? entry.depth || 0 : 0;
        const max = isObject ? entry.ancestors?.length || 0 : 0;
        const depthHtml =
            max > 0
                ? `<span style="display:inline-flex;align-items:center;gap:2px;">
                    <button type="button" data-rule-depth="-1" data-rule-target="${role}:${index}" style="${depthBtnStyle}">−</button>
                    <span style="min-width:12px;text-align:center;">${depth}</span>
                    <button type="button" data-rule-depth="1" data-rule-target="${role}:${index}" style="${depthBtnStyle}">+</button>
                   </span>`
                : `<span style="font-size:10px;color:#9ca3af;min-width:54px;text-align:center;">—</span>`;

        return `
            <div style="display:flex;align-items:center;gap:4px;margin-bottom:4px;font-size:11px;">
                <span style="color:#9ca3af;min-width:14px;">${index + 1}.</span>
                <code style="flex:1;word-break:break-all;color:#374151;" title="${resolved.replace(/"/g, "&quot;")}">${shortSelector(resolved)}</code>
                ${depthHtml}
                <button type="button" data-repick-index="${role}:${index}" title="Re-pick" style="border:none;background:#e5e7eb;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:11px;">↻</button>
                <button type="button" data-remove-rule="${role}:${index}" title="Remove" style="border:none;background:#fecaca;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:11px;color:#991b1b;">×</button>
            </div>
        `;
    }

    function renderMappingRules() {
        if (!mappingRulesEl) return;

        mappingRulesEl.innerHTML = Object.keys(ROLE_LABELS)
            .map(role => {
                const rules = getRoleList(role);
                const rows = rules.length
                    ? rules.map((entry, i) => renderRuleRow(role, entry, i)).join("")
                    : `<div style="font-size:11px;color:#9ca3af;margin-bottom:4px;">No rules yet</div>`;

                return `
                    <div style="margin-bottom:12px;">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                            <span style="font-weight:600;font-size:13px;">${ROLE_LABELS[role]}</span>
                            <button type="button" data-add="${role}" style="border:none;background:#dbeafe;color:#1d4ed8;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:11px;font-weight:600;">+ Add</button>
                        </div>
                        ${rows}
                    </div>
                `;
            })
            .join("");
    }

    function refreshUi() {
        if (!panel) return;

        const ready = config.configured;
        setupSection.style.display = ready ? "none" : "block";
        themeSection.style.display = ready ? "block" : "none";
        mappingSection.style.display = ready ? "block" : "none";
        focusToggle.style.display =
            config.selectors.hide.length ? "block" : "none";

        renderMappingRules();

        bgHue.value = config.colors.bgHue;
        bgLight.value = config.colors.bgLightness;
        textLight.value = config.colors.textLightness;
        bgHueValue.textContent = config.colors.bgHue;
        bgLightValue.textContent = config.colors.bgLightness;
        textLightValue.textContent = config.colors.textLightness;
    }

    const container = document.createElement("div");
    container.id = ROOT_ID;
    container.style.cssText = `
        position:fixed; bottom:20px; left:20px; z-index:2147483645;
        font-family:Inter,Arial,sans-serif;
    `;

    const toggleBtn = document.createElement("div");
    toggleBtn.textContent = "Aa Reader";
    toggleBtn.style.cssText = `
        background:#111827; color:#fff; padding:12px 18px; border-radius:999px;
        cursor:pointer; font-weight:600; user-select:none;
        box-shadow:0 8px 24px rgba(0,0,0,.25);
    `;

    const panel = document.createElement("div");
    panel.style.cssText = `
        position:absolute; bottom:60px; left:0; width:340px;
        background:#fff; color:#111827; border-radius:16px; padding:16px;
        display:none; box-shadow:0 12px 40px rgba(0,0,0,.25);
        max-height:70vh; overflow:auto;
    `;

    const btnStyle = `
        width:100%; border:none; border-radius:10px; padding:10px;
        cursor:pointer; color:#fff; background:#374151; font-weight:600;
        margin-bottom:8px; text-align:left;
    `;

    panel.innerHTML = `
        <div style="font-size:16px;font-weight:700;margin-bottom:4px;">
            Reader Settings
        </div>
        <div style="font-size:12px;color:#6b7280;margin-bottom:12px;" id="grt-host"></div>

        <div id="grt-setup">
            <div style="font-size:13px;margin-bottom:10px;line-height:1.4;">
                Set up this site once. Pick the reading area and text, then tune colors.
            </div>
            <button type="button" data-pick="page" style="${btnStyle}background:#1d4ed8;">
                Add page background
            </button>
            <button type="button" data-pick="panel" style="${btnStyle}background:#1d4ed8;">
                Add content panel
            </button>
            <button type="button" data-pick="text" style="${btnStyle}background:#1d4ed8;">
                Add text target
            </button>
            <button type="button" data-pick="hide" style="${btnStyle}background:#6b7280;">
                Add element to hide (optional)
            </button>
        </div>

        <div id="grt-theme" style="display:none;">
            <button type="button" id="grt-focus-toggle" style="${btnStyle}background:#111827;margin-bottom:16px;">
                Zen Mode: Off
            </button>

            <div style="margin-bottom:14px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                    <span>Background hue</span>
                    <span id="grt-bgHueValue"></span>
                </div>
                <input id="grt-bgHue" type="range" min="0" max="360" style="width:100%;">
            </div>

            <div style="margin-bottom:14px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                    <span>Background brightness</span>
                    <span id="grt-bgLightValue"></span>
                </div>
                <input id="grt-bgLight" type="range" min="0" max="100" style="width:100%;">
            </div>

            <div style="margin-bottom:14px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                    <span>Text brightness</span>
                    <span id="grt-textLightValue"></span>
                </div>
                <input id="grt-textLight" type="range" min="0" max="100" style="width:100%;">
            </div>
        </div>

        <div id="grt-mapping" style="display:none;margin-top:8px;padding-top:12px;border-top:1px solid #e5e7eb;">
            <div style="font-size:13px;font-weight:700;margin-bottom:4px;">Site mapping</div>
            <div style="font-size:11px;color:#6b7280;margin-bottom:10px;line-height:1.4;">
                Add multiple rules per type. Use +/− on each rule to beat site CSS specificity.
            </div>
            <div id="grt-mapping-rules"></div>
            <button type="button" id="grt-reset" style="${btnStyle}background:#b91c1c;margin-top:8px;">
                Reset site settings
            </button>
        </div>
    `;

    container.appendChild(panel);
    container.appendChild(toggleBtn);

    const setupSection = panel.querySelector("#grt-setup");
    const themeSection = panel.querySelector("#grt-theme");
    const mappingSection = panel.querySelector("#grt-mapping");
    const mappingRulesEl = panel.querySelector("#grt-mapping-rules");
    const focusToggle = panel.querySelector("#grt-focus-toggle");
    const bgHue = panel.querySelector("#grt-bgHue");
    const bgLight = panel.querySelector("#grt-bgLight");
    const textLight = panel.querySelector("#grt-textLight");
    const bgHueValue = panel.querySelector("#grt-bgHueValue");
    const bgLightValue = panel.querySelector("#grt-bgLightValue");
    const textLightValue = panel.querySelector("#grt-textLightValue");

    panel.querySelector("#grt-host").textContent = host;

    const pickLabels = {
        page: "page background",
        panel: "content panel",
        text: "body text",
        hide: "element to hide in zen mode"
    };

    function openPicker(role, options) {
        panel.style.display = "none";
        startPicker(role, pickLabels[role], options || { mode: "append" }, () => {
            panel.style.display = "block";
        });
    }

    function bindPick(button, role) {
        button.addEventListener("click", e => {
            e.stopPropagation();
            openPicker(role, { mode: "append" });
        });
    }

    panel.querySelectorAll("[data-pick]").forEach(btn => {
        bindPick(btn, btn.dataset.pick);
    });

    mappingSection.addEventListener("click", e => {
        const addBtn = e.target.closest("[data-add]");
        if (addBtn) {
            e.stopPropagation();
            openPicker(addBtn.dataset.add, { mode: "append" });
            return;
        }

        const repickBtn = e.target.closest("[data-repick-index]");
        if (repickBtn) {
            e.stopPropagation();
            const [role, index] = repickBtn.dataset.repickIndex.split(":");
            openPicker(role, { mode: "replace", index: Number(index) });
            return;
        }

        const removeBtn = e.target.closest("[data-remove-rule]");
        if (removeBtn) {
            e.stopPropagation();
            const [role, index] = removeBtn.dataset.removeRule.split(":");
            removeRule(role, Number(index));
            applyTheme();
            return;
        }

        const depthBtn = e.target.closest("[data-rule-depth]");
        if (depthBtn) {
            e.stopPropagation();
            const [role, index] = depthBtn.dataset.ruleTarget.split(":");
            setRuleDepth(role, Number(index), Number(depthBtn.dataset.ruleDepth));
        }
    });

    bgHue.addEventListener("input", () => {
        config.colors.bgHue = Number(bgHue.value);
        applyTheme();
    });

    bgLight.addEventListener("input", () => {
        config.colors.bgLightness = Number(bgLight.value);
        applyTheme();
    });

    textLight.addEventListener("input", () => {
        config.colors.textLightness = Number(textLight.value);
        applyTheme();
    });

    focusToggle.addEventListener("click", e => {
        e.stopPropagation();
        setFocusMode(!document.body.classList.contains(FOCUS_CLASS));
    });

    panel.querySelector("#grt-reset").addEventListener("click", e => {
        e.stopPropagation();
        if (!confirm(`Reset reader settings for ${host}?`)) return;
        localStorage.removeItem(storageKey);
        localStorage.removeItem(focusKey);
        config = loadConfig();
        if (!PRESETS[host]) {
            config = emptyConfig();
        }
        setFocusMode(false);
        applyTheme();
    });

    toggleBtn.addEventListener("click", e => {
        e.stopPropagation();
        panel.style.display = panel.style.display === "none" ? "block" : "none";
    });

    panel.addEventListener("click", e => e.stopPropagation());

    document.addEventListener("click", () => {
        panel.style.display = "none";
    });

    function mountUi() {
        if (document.getElementById(ROOT_ID)) return;
        (document.body || document.documentElement).appendChild(container);
    }

    if (document.body) {
        mountUi();
    } else {
        document.addEventListener("DOMContentLoaded", mountUi, { once: true });
    }

    applyTheme();

    function initFocusMode() {
        setFocusMode(localStorage.getItem(focusKey) === "true");
    }

    if (document.body) {
        initFocusMode();
    } else {
        document.addEventListener("DOMContentLoaded", initFocusMode, { once: true });
    }

    let observerTimer = null;
    const observer = new MutationObserver(() => {
        clearTimeout(observerTimer);
        observerTimer = setTimeout(applyTheme, 300);
    });

    if (document.body) {
        observer.observe(document.body, { childList: true, subtree: true });
    } else {
        document.addEventListener(
            "DOMContentLoaded",
            () => observer.observe(document.body, { childList: true, subtree: true }),
            { once: true }
        );
    }
})();
