// ==UserScript==
// @name        Webnovel filter novels
// @namespace   Violentmonkey Scripts
// @version     1.0.0
//
// @match       https://www.webnovel.com/stories/*
// @match       https://www.webnovel.com/stories
// @grant       none
//
// @author      -
// @description
// ==/UserScript==
(function () {
    if (window.__novelFilterInstalled) return;
    window.__novelFilterInstalled = true;

    let minRating = 4.7;
    let minChapters = 0;
    let filterEnabled = true;

    const tagFilters = new Map([
        ["R18", "exclude"],
        ["BL", "exclude"],
        ["R18", "exclude"],
        ["R18", "exclude"],
    ]);

    function getRating(li) {
        const spans = li.querySelectorAll("span.vam.fw400.lh16.fs12");

        for (const span of spans) {
            const value = parseFloat(span.textContent.trim());

            if (!isNaN(value) && value <= 5) {
                return value;
            }
        }

        return null;
    }

    function getTags(li) {
        return [...li.querySelectorAll('a[href*="/tags/"]')]
            .map(a =>
                a.textContent
                    .replace(/^#/, "")
                    .trim()
                    .toUpperCase()
            );
    }

    function getChapters(li) {
        const spans = li.querySelectorAll("span.vam.fw400.lh16.fs12");

        for (const span of spans) {
            const match =
                span.textContent
                    .trim()
                    .match(/^([\d,]+)\s+Chapters?$/i);

            if (match) {
                return parseInt(
                    match[1].replace(/,/g, ""),
                    10
                );
            }
        }

        return null;
    }

    function passesTagFilter(tags) {
        for (const [tag, mode] of tagFilters) {
            if (mode === "include" && !tags.includes(tag)) {
                return false;
            }

            if (mode === "exclude" && tags.includes(tag)) {
                return false;
            }
        }

        return true;
    }

    function applyFilter() {
        document.querySelectorAll("li.fl").forEach(li => {
            const rating = getRating(li);
            const chapters = getChapters(li);
            const tags = getTags(li);

            const ratingPass =
                !filterEnabled ||
                rating === null ||
                rating >= minRating;

            const tagPass = passesTagFilter(tags);

            const chaptersPass =
                chapters === null ||
                chapters >= minChapters;

            li.style.display =
                ratingPass && tagPass && chaptersPass
                    ? ""
                    : "none";
        });

        updateCounter();
    }

    function updateCounter() {
        const all = document.querySelectorAll("li.fl").length;

        const visible =
            [...document.querySelectorAll("li.fl")]
                .filter(li => li.style.display !== "none")
                .length;

        counter.textContent =
            `Showing ${visible} / ${all}`;
    }

    function cycleTagState(btn) {
        const tag = btn.dataset.tag;
        const state = btn.dataset.state;

        if (state === "neutral") {
            btn.dataset.state = "include";
            btn.style.background = "#16a34a";
            btn.style.color = "white";
            tagFilters.set(tag, "include");
        }
        else if (state === "include") {
            btn.dataset.state = "exclude";
            btn.style.background = "#dc2626";
            btn.style.color = "white";
            tagFilters.set(tag, "exclude");
        }
        else {
            btn.dataset.state = "neutral";
            btn.style.background = "#e5e7eb";
            btn.style.color = "#111827";
            tagFilters.delete(tag);
        }
    }

    function styleTagButton(btn, state) {
        if (state === "include") {
            btn.style.background = "#16a34a";
            btn.style.color = "white";
        }
        else if (state === "exclude") {
            btn.style.background = "#dc2626";
            btn.style.color = "white";
        }
        else {
            btn.style.background = "#e5e7eb";
            btn.style.color = "#111827";
        }
    }

    function renderTagButtons(tagCounts) {
        tagContainer.replaceChildren();

        [...tagCounts.entries()]
            .sort(([tagA, countA], [tagB, countB]) =>
                countB - countA ||
                tagA.localeCompare(tagB)
            )
            .forEach(([tag, count]) => {
                const btn = document.createElement("span");
                const state = tagFilters.get(tag) || "neutral";

                btn.dataset.tag = tag;
                btn.dataset.state = state;
                btn.textContent = `${tag} (${count})`;

                btn.style.cssText = `
                    display:inline-block;
                    margin:3px;
                    padding:4px 10px;
                    border-radius:999px;
                    cursor:pointer;
                    background:#e5e7eb;
                    color:#111827;
                    font-size:11px;
                    user-select:none;
                    transition:all .15s ease;
                `;

                styleTagButton(btn, state);

                btn.addEventListener("click", () => {
                    cycleTagState(btn);
                    applyFilter();
                });

                tagContainer.appendChild(btn);
            });
    }

    function refreshTagList() {
        const tagCounts = new Map();

        document
            .querySelectorAll("li.fl")
            .forEach(li => {
                getTags(li).forEach(tag => {
                    tagCounts.set(
                        tag,
                        (tagCounts.get(tag) || 0) + 1
                    );
                });
            });

        renderTagButtons(tagCounts);
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

    button.textContent = "⭐ Filter";

    button.style.cssText = `
        background:#111827;
        color:white;
        padding:12px 18px;
        border-radius:999px;
        cursor:pointer;
        font-weight:600;
        box-shadow:0 8px 24px rgba(0,0,0,.25);
        user-select:none;
    `;

    const panel = document.createElement("div");

    panel.style.cssText = `
        position:absolute;
        bottom:60px;
        left:0;
        width:380px;
        max-height:600px;
        overflow:auto;
        background:white;
        border-radius:16px;
        padding:16px;
        box-shadow:0 12px 40px rgba(0,0,0,.25);
        display:none;
    `;

    panel.innerHTML = `
        <div style="
            font-size:16px;
            font-weight:700;
            margin-bottom:12px;
        ">
            Novel Filters
        </div>

        <div style="margin-bottom:12px;">
            <label style="
                display:block;
                margin-bottom:6px;
                font-size:12px;
            ">
                Minimum Rating
            </label>

            <input
                id="novel-filter-rating"
                type="number"
                min="0"
                max="5"
                step="0.1"
                value="${minRating}"
                style="
                    width:100%;
                    padding:8px;
                    border:1px solid #ddd;
                    border-radius:8px;
                    box-sizing:border-box;
                "
            />
        </div>

        <div style="margin-bottom:12px;">
            <label style="
                display:block;
                margin-bottom:6px;
                font-size:12px;
            ">
                Minimum Chapters
            </label>

            <input
                id="novel-filter-chapters"
                type="number"
                min="0"
                step="1"
                value="${minChapters}"
                style="
                    width:100%;
                    padding:8px;
                    border:1px solid #ddd;
                    border-radius:8px;
                    box-sizing:border-box;
                "
            />
        </div>

        <label style="
            display:flex;
            align-items:center;
            gap:8px;
            margin-bottom:12px;
        ">
            <input
                id="novel-filter-enabled"
                type="checkbox"
                checked
            />
            Rating Filter Enabled
        </label>

        <div id="novel-counter"
             style="
                margin-bottom:12px;
                font-size:12px;
                color:#666;
             ">
        </div>

        <input
            id="tag-search"
            type="text"
            placeholder="Search tags..."
            style="
                width:100%;
                box-sizing:border-box;
                padding:8px;
                border:1px solid #ddd;
                border-radius:8px;
                margin-bottom:10px;
            "
        />

        <div style="
            font-size:12px;
            color:#666;
            margin-bottom:8px;
        ">
            Gray = Ignore • Green = Require • Red = Exclude
        </div>

        <div
            id="novel-tag-container"
            style="
                max-height:220px;
                overflow-y:auto;
                padding-right:4px;
            "
        ></div>
    `;

    container.appendChild(panel);
    container.appendChild(button);
    document.body.appendChild(container);

    const tagContainer =
        document.getElementById("novel-tag-container");

    const counter =
        document.getElementById("novel-counter");

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

    const ratingInput =
        document.getElementById("novel-filter-rating");

    const chaptersInput =
        document.getElementById("novel-filter-chapters");

    const enabledInput =
        document.getElementById("novel-filter-enabled");

    function updateFilterSettings() {
        const nextMinRating = parseFloat(ratingInput.value);
        const nextMinChapters = parseInt(chaptersInput.value, 10);

        if (!isNaN(nextMinRating)) {
            minRating = nextMinRating;
        }

        minChapters =
            !isNaN(nextMinChapters)
                ? nextMinChapters
                : 0;

        filterEnabled = enabledInput.checked;
        applyFilter();
    }

    ratingInput.addEventListener("input", updateFilterSettings);
    chaptersInput.addEventListener("input", updateFilterSettings);
    enabledInput.addEventListener("change", updateFilterSettings);

    document
        .getElementById("tag-search")
        .addEventListener("input", function () {
            const search =
                this.value
                    .trim()
                    .toUpperCase();

            [...tagContainer.children]
                .forEach(btn => {
                    btn.style.display =
                        btn.dataset.tag.includes(search)
                            ? "inline-block"
                            : "none";
                });
        });

    let observerScheduled = false;

    function nodeContainsNovelItem(node) {
        if (!(node instanceof Element)) {
            return false;
        }

        if (node === container || container.contains(node)) {
            return false;
        }

        return (
            node.matches?.("li.fl") ||
            node.querySelector?.("li.fl")
        );
    }

    const observer =
        new MutationObserver(mutations => {
            window.dispatchEvent(new Event("scroll"));

            const hasNewNovelItems = mutations.some(mutation =>
                [...mutation.addedNodes].some(nodeContainsNovelItem)
            );

            if (!hasNewNovelItems || observerScheduled) {
                return;
            }

            observerScheduled = true;

            requestAnimationFrame(() => {
                observerScheduled = false;
                refreshTagList();
                applyFilter();
            });
        });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

    refreshTagList();
    applyFilter();
})();
