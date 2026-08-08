// ==UserScript==
// @name        hover over word
// @namespace   Violentmonkey Scripts
// @match       https://www.webnovel.com/book/*/*
// @require     https://cdn.jsdelivr.net/npm/@violentmonkey/dom@2
// @grant       none
// @version     1.0
// @author      bean
// @run-at      document-body
// @description 14.3.2024, 22:57:41
// ==/UserScript==

const oldContentAttribute = "old-content-data";

const mutationCallback = function (mutationsList, observer) {
    for (let mutation of mutationsList) {
        if (mutation.type === "childList") {
            // Enable text selection
            document.body.style.userSelect = 'auto';

            // Enable right-click
            document.oncontextmenu = null;
            // New nodes have been added or removed
            mutation.addedNodes.forEach((node) => {
                if (!node.classList?.contains('chapter_content')) return;
                const paragraphs = node.querySelectorAll('div.cha-content div.cha-paragraph div p');
                paragraphs.forEach(p => {
                    p.classList.add("hoverable");
                    const words = p.textContent.split(" ");
                    p.setAttribute(oldContentAttribute, p.textContent);
                    p.textContent = "";
                    words.forEach((word,index) => {
                        const span = document.createElement("span");
                        span.textContent = word + (index + 1 < words.length ? " " : "");
                        p.appendChild(span);
                    });
                })

            });
        }
    }
};

const observer = new MutationObserver(mutationCallback);

// Start observing the DOM for changes
observer.observe(document.body, {
    childList: true,
    subtree: true,
});

var css = `
p.hoverable span:hover{
    color: lightcoral !important;
}

* {
    user-select: auto !important;
}
::selection {
    background-color: transparent !important;
}

::-moz-selection {
    background-color: transparent !important;
}

p.hoverable span::selection {
    color: deepskyblue !important;
}
  `;

var style = document.createElement("style");
style.innerText = css;
document.head.appendChild(style);
