// ==UserScript==
// @name        Webnovel Reader Custom Theme 2
// @namespace   Violentmonkey Scripts
// @version     2.0.0
//
// @match       https://www.webnovel.com/book/*/*
// @grant       none
//
// @author      -
// @description Custom reader theme with color sliders
// ==/UserScript==

(function () {
    if (window.__novelReaderInstalled) return;
    window.__novelReaderInstalled = true;

    const SETTINGS_KEY = "novelReaderSettings";
    const FOCUS_KEY = "novelReaderFocusMode";

    const defaults = {
        bgHue: 35,
        bgLightness: 92,
        textLightness: 15
    };

    const settings = {
        ...defaults,
        ...(JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}"))
    };

    const style = document.createElement("style");
    document.head.appendChild(style);

    function saveSettings() {
        localStorage.setItem(
            SETTINGS_KEY,
            JSON.stringify(settings)
        );
    }

    function hsl(h, s, l) {
        return `hsl(${h}, ${s}%, ${l}%)`;
    }

    function clamp(v, min, max) {
        return Math.max(min, Math.min(max, v));
    }

    function getTheme() {
        const page = hsl(
            settings.bgHue,
            35,
            settings.bgLightness
        );

        const panel = hsl(
            settings.bgHue,
            30,
            clamp(settings.bgLightness + 4, 0, 100)
        );

        const border = hsl(
            settings.bgHue,
            25,
            clamp(settings.bgLightness - 10, 0, 100)
        );

        const text = hsl(
            settings.bgHue,
            20,
            settings.textLightness
        );

        const muted = hsl(
            settings.bgHue,
            15,
            clamp(settings.textLightness + 25, 0, 100)
        );

        const accent = hsl(
            settings.bgHue,
            60,
            45
        );

        return {
            page,
            panel,
            border,
            text,
            muted,
            accent
        };
    }

    function focusCss() {
        return `
            body.__novel-reader-focus {
                overflow-x:hidden !important;
                overflow-y:auto !important;
            }

            body.__novel-reader-focus header.cha-header,
            body.__novel-reader-focus div.cha-fly {
                display:none !important;
            }
        `;
    }

    function applyTheme() {
        const t = getTheme();

        style.textContent = `
            body {
                --bc_dark_primary:${t.panel} !important;
            }

            html,
            body
             {
                background:${t.page} !important;
            }

            .cha-page,
            ._color3 .cha-page {
                background:${t.panel} !important;
                color:${t.text} !important;
                border-color:${t.border} !important;
            }

            .cha-page {
                border:1px solid ${t.border} !important;
            }

            ._color3 .cha-content,
            .cha-content,
            .cha-content p,
            .cha-content span {
                color:${t.text} !important;
            }

            .cha-page * {
                border-color:${t.border} !important;
            }

            a {
                color:${t.accent} !important;
            }

            ${focusCss()}
        `;

        bgHueValue.textContent =
            settings.bgHue;

        bgLightValue.textContent =
            settings.bgLightness;

        textLightValue.textContent =
            settings.textLightness;

        saveSettings();
    }

    function setFocusMode(enabled) {
        document.body.classList.toggle(
            "__novel-reader-focus",
            enabled
        );

        localStorage.setItem(
            FOCUS_KEY,
            enabled ? "true" : "false"
        );

        focusToggle.textContent =
            enabled
                ? "Zen Mode: On"
                : "Zen Mode: Off";

        focusToggle.style.background =
            enabled
                ? "#16a34a"
                : "#111827";
    }

    const container = document.createElement("div");

    container.style.cssText = `
        position:fixed;
        bottom:20px;
        left:20px;
        z-index:999999;
        font-family:Inter,Arial,sans-serif;
    `;

    const button = document.createElement("div");

    button.textContent = "Aa Reader";

    button.style.cssText = `
        background:#111827;
        color:white;
        padding:12px 18px;
        border-radius:999px;
        cursor:pointer;
        font-weight:600;
        user-select:none;
        box-shadow:0 8px 24px rgba(0,0,0,.25);
    `;

    const panel = document.createElement("div");

    panel.style.cssText = `
        position:absolute;
        bottom:60px;
        left:0;
        width:320px;
        background:white;
        color:#111827;
        border-radius:16px;
        padding:16px;
        display:none;
        box-shadow:0 12px 40px rgba(0,0,0,.25);
    `;

    panel.innerHTML = `
        <div style="
            font-size:16px;
            font-weight:700;
            margin-bottom:12px;
        ">
            Reader Settings
        </div>

        <button
            id="novel-focus-toggle"
            style="
                width:100%;
                border:none;
                border-radius:10px;
                padding:10px;
                cursor:pointer;
                color:white;
                background:#111827;
                font-weight:700;
                margin-bottom:16px;
            "
        >
            Zen Mode: Off
        </button>

        <div style="margin-bottom:14px;">
            <div style="
                display:flex;
                justify-content:space-between;
                margin-bottom:4px;
            ">
                <span>Background Color</span>
                <span id="bgHueValue"></span>
            </div>

            <input
                id="bgHue"
                type="range"
                min="0"
                max="360"
                value="${settings.bgHue}"
                style="width:100%;"
            >
        </div>

        <div style="margin-bottom:14px;">
            <div style="
                display:flex;
                justify-content:space-between;
                margin-bottom:4px;
            ">
                <span>Background Brightness</span>
                <span id="bgLightValue"></span>
            </div>

            <input
                id="bgLight"
                type="range"
                min="0"
                max="100"
                value="${settings.bgLightness}"
                style="width:100%;"
            >
        </div>

        <div>
            <div style="
                display:flex;
                justify-content:space-between;
                margin-bottom:4px;
            ">
                <span>Text Brightness</span>
                <span id="textLightValue"></span>
            </div>

            <input
                id="textLight"
                type="range"
                min="0"
                max="100"
                value="${settings.textLightness}"
                style="width:100%;"
            >
        </div>
    `;

    container.appendChild(panel);
    container.appendChild(button);

    document.body.appendChild(container);

    const focusToggle =
        panel.querySelector("#novel-focus-toggle");

    const bgHue =
        panel.querySelector("#bgHue");

    const bgLight =
        panel.querySelector("#bgLight");

    const textLight =
        panel.querySelector("#textLight");

    const bgHueValue =
        panel.querySelector("#bgHueValue");

    const bgLightValue =
        panel.querySelector("#bgLightValue");

    const textLightValue =
        panel.querySelector("#textLightValue");

    bgHue.addEventListener("input", () => {
        settings.bgHue = Number(bgHue.value);
        applyTheme();
    });

    bgLight.addEventListener("input", () => {
        settings.bgLightness =
            Number(bgLight.value);

        applyTheme();
    });

    textLight.addEventListener("input", () => {
        settings.textLightness =
            Number(textLight.value);

        applyTheme();
    });

    focusToggle.addEventListener("click", () => {
        setFocusMode(
            !document.body.classList.contains(
                "__novel-reader-focus"
            )
        );
    });

    button.addEventListener("click", e => {
        e.stopPropagation();

        panel.style.display =
            panel.style.display === "none"
                ? "block"
                : "none";
    });

    panel.addEventListener("click", e => {
        e.stopPropagation();
    });

    document.addEventListener("click", () => {
        panel.style.display = "none";
    });

    applyTheme();

    setFocusMode(
        localStorage.getItem(FOCUS_KEY) === "true"
    );
})();
