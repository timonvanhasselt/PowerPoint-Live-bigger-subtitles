// ==UserScript==
// @name         PowerPoint Live Subtitle Enlarger
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Enlarges the subtitles in PowerPoint Live and hides specific elements
// @author       You
// @match        https://euc-powerpoint.officeapps.live.com/*
// @match        https://oauth.officeapps.live.com/*
// @match        https://webshell.suite.office.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    function hideElements() {
        const elementsToHide = [
            "slideshowBeginningFocusTrap",
            "SlideShowContainer",
            "HiddenFocusTrapAnchor",
            "slideshowEndFocusTrap",
            "screenReaderDetectionDiv",
            "startingSlideshowHiddenDiv"
        ];

        elementsToHide.forEach(id => {
            let element = document.getElementById(id);
            if (element) {
                element.style.display = "none";
            }
        });
    }

    function enlargeSubtitles() {
        let subtitleSpan = document.getElementById("SubtitleResultSpan");
        let subtitleContainer = document.getElementById("SubtitleContainer");

        if (subtitleSpan) {
            subtitleSpan.style.fontSize = "100px";
            subtitleSpan.style.lineHeight = "120px"; // Prevent overlap
            subtitleSpan.style.padding = "10px";
            subtitleSpan.style.color = "white"; // For good contrast
        }

        if (subtitleContainer) {
            subtitleContainer.style.height = "90%"; // Set the height to 90%
        }
    }

    function detectSubtitleWindow() {
        if (document.getElementById("SubtitleResultSpan")) {
            console.log("✅ Subtitle window detected, script active.");
            enlargeSubtitles();
            hideElements(); // Hide the elements

            // Observer to automatically adjust changes
            const observer = new MutationObserver(() => {
                enlargeSubtitles();
                hideElements(); // Hide again if there are changes
            });
            observer.observe(document.body, { childList: true, subtree: true });
        } else {
            console.log("❌ No subtitle window found. Continuing to search...");
            setTimeout(detectSubtitleWindow, 500); // Check every half second
        }
    }

    detectSubtitleWindow(); // Start the detection
})();
