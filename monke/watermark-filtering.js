// ==UserScript==
// @name        filter dogshit watermark v2 novel1st.com
// @namespace   Violentmonkey Scripts
// @match       https://novel1st.com/*
// @grant       none
// @require     https://cdn.jsdelivr.net/npm/@violentmonkey/dom@2
// @version     1.0
// @author      -
// @description 4.3.2025, 20:54:59
// ==/UserScript==


const getParagraphs = () => {
  return document.querySelectorAll('.reading-content p');
};

VM.observe(document.body, () => {
  const peas = getParagraphs();
  if (peas.length === 0) return
  console.log('replacing');
  peas.forEach(p => {
    if (p.textContent.match(/Follow new episodes on the.*/,'')) p.textContent = '';
  });
  return true;

})
