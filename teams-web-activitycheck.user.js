// ==UserScript==
// @name         Teams Web - Activity Check with Idle Detector
// @namespace    http://tampermonkey.net/
// @version      2024-01-10_01-00
// @description  Makes Teams' activity check use the Idle Detector API so activity is tracked not only within the tab, but system-wide.
// @author       tsukasa
// @homepage     https://github.com/tsukasa/teams-web-activitycheck
// @updateURL    https://raw.githubusercontent.com/tsukasa/teams-web-activitycheck/master/teams-web-activitycheck.user.js
// @downloadURL  https://raw.githubusercontent.com/tsukasa/teams-web-activitycheck/master/teams-web-activitycheck.user.js
// @match        https://teams.microsoft.com/*
// @icon         data:image/png;base64,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
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function() {
    'use strict';

    /***********************************************************************************************
     * HOW TO USE:                                                                                 *
     * 1. Install and enable the script.                                                           *
     * 2. Reload the Teams web client.                                                             *
     * 3. Now click on the Teams icon in the top left of Teams.                                    *
     * 4. A popup that asks for Idle Detector permissions should pop up. Grant the permissions.    *
     * 5. Congratulations, from now on it should just work.                                        *
     ***********************************************************************************************/

    const _intervalMin = 1 * 60 * 1000;
    const _intervalMax = 2 * 60 * 1000;

    const _titlebarIconSelector = "svg[data-tid='titlebar-teams-icon']";

    /**
     * Mechanisms:
     * 0  setTimeout-based loop that also works without the IdleDetector API.
     * 1  IdleDetector that should be the preferred solution (DEFAULT).
     *
     * Defaults to 1 if an invalid number is given.
     */
    let _mechanism = 1;

    let _performLogging = false;
    let _idleDetector = null;
    let _timeoutId = null;

    /**
     * Logs a message to the console with a specific format.
     *
     * @param {string} type    - The type of console log method to use (e.g., 'log', 'warn', 'error').
     * @param {...any} message - The message(s) to be logged.
     */
    const log = (type, ...message) => {
        if (!_performLogging) return;
        window.__tm_console[type]("%c[tsu.teams]%c", "color: #3E82E5; font-weight: 700;", "", ...message);
    };

    /**
     * Restore the native console functions by injecting an iframe.
     * Required because Teams' console functions are overridden.
     * Sourced from https://dev.to/js_bits_bill/how-to-restore-native-browser-code-3e6e
     */
    const createConsoleProxy = () => {
        // Create dummy iframe to steal its fresh console object
        const iframe = document.createElement("iframe");

        // Add iframe to current window's scope in a hidden state
        iframe.id = "console-proxy";
        iframe.style.display = "none";
        document.body.insertAdjacentElement("beforeend", iframe);

        // Reassign value of console to iframe's console
        const proxyIframe = document.querySelector("#console-proxy");
        window.__tm_console = proxyIframe.contentWindow.console;

        if("__tm_console" in window) {
            _performLogging = true;
        }
    };

    /**
     * Checks whether a given configuration value for `_mechanism` is valid.
     * Defaults the value to 1 if invalid value was found.
     */
    const checkIfMechanismValid = () => {
        if (_mechanism !== 0 && _mechanism !== 1) {
            _mechanism = 1;
        }
    };

    /**
     * Event handler for the Teams icon in the titlebar.
     * Used to request IdleDetector permissions.
     */
    const titlebarTeamsIconClickHandler = async () => {
        log("log", "Requesting IdleDetector permissions...");
        const res = await window.IdleDetector.requestPermission();

        if (res !== "granted") {
            throw new Error("IdleDetector permissions not granted.");
        }

        await startIdleDetector();
    };

    /**
     * Adds or removes the event handler for the Teams icon in the titlebar.
     * @param enable - Whether to enable or disable the event handler.
     */
    const configureTitlebarTeamsIconClickHandler = (enable) => {
        const loadingObserver = new MutationObserver(() => {
            const appIcon = document.querySelector(_titlebarIconSelector);

            if (!appIcon) return;

            loadingObserver.disconnect();

            if (enable) {
                appIcon.addEventListener("click", titlebarTeamsIconClickHandler);
                log("debug", "Added titlebar-teams-icon click event handler.");
            } else {
                appIcon.removeEventListener("click", titlebarTeamsIconClickHandler);
                log("debug", "Removed titlebar-teams-icon click event handler.");
            }
        });

        loadingObserver.observe(document, {
            childList: true,
            subtree: true
        });
    }

    /**
     * Starts the IdleDetector API.
     */
    const startIdleDetector = async () => {
        try {
            log("debug", "Starting IdleDetector...");

            await _idleDetector.start({
                threshold: _intervalMax,
                signal: (new AbortController()).signal,
            });

            configureTitlebarTeamsIconClickHandler(false);

            log("log", "IdleDetector started.");
        } catch (err) {
            log("error", err.name, err.message);
        }
    };

    /**
     * Sets up the IdleDetector API and starts it.
     * @returns {Promise<void>}
     */
    const setupIdleDetector = async () => {
        // Make sure the IdleDetector API is even available...
        if (!"IdleDetector" in window) {
            log("debug", "IdleDetector API not available, falling back to regular looping.");
            _mechanism = 0;
            return;
        }

        try {
            configureTitlebarTeamsIconClickHandler(true);

            _idleDetector = new window.IdleDetector();
            _idleDetector.addEventListener('change', () => {
                log("info", `IdleDetector state change: userState = ${_idleDetector.userState}, screenState = ${_idleDetector.screenState}`);
                // If user is inactive, cancel a scheduled run and exit.
                if (_idleDetector.userState !== "active" || _idleDetector.screenState !== "unlocked") {
                    clearTimeout(_timeoutId);
                    return;
                };
                // If we have come this far, the operation should be performed immediatly since the loop has been broken.
                // doExecute does queue up the necessary calls to re-enable the loop.
                doExecute();
            });
            await startIdleDetector();
        } catch (err) {
            log("error", err.name, err.message);
        }
    }

    /**
     * Grouped helper function that performs the relevant actions for dispatching and looping.
     */
    const doExecute = () => {
        dispatchMouseEvent();
        prepareExecute();
    };

    /**
     * Helper function that checks whether the user is active or not based on the IdleDetector API.
     * @returns {boolean} Whether the user is active or not.
     */
    const checkIfUserIsActive = () => {
        if (_idleDetector.userState !== "active") {
            log("debug", `According to IdleDetector, user is ${_idleDetector.userState}.`);
            return false;
        }

        if (_idleDetector.screenState !== "unlocked") {
            log("debug", `According to IdleDetector, screen is ${_idleDetector.screenState}.`);
            return false;
        }

        return true;
    }

    /**
     * Queues up the next execution of the loop and generates a random wait time.
     */
    const prepareExecute = () => {
        // When using the IdleDetector API, only fire the event if user is active.
        if (_mechanism == 1 && checkIfUserIsActive() === false) return;

        const waitTime = generateRandomNumber(_intervalMin, _intervalMax);
        _timeoutId = setTimeout(doExecute, waitTime);
        log("log", `Next execution in ${waitTime}ms.`);
    };

    /**
     * Helper function to generate a random number between two given numbers.
     * @param {number} min - The minimum number to generate.
     * @param {number} max - The maximum number to generate.
     * @returns {number} - A random number between min and max.
     */
    const generateRandomNumber = (min, max) => {
        return Math.floor(Math.random() * (max - min + 1) + min);
    };

    /**
     * Helper function to dispatch a mousemove event to the document.
     */
    const dispatchMouseEvent = () => {
        const x = generateRandomNumber(1, window.innerWidth / 2);
        const y = generateRandomNumber(1, window.innerHeight / 2);

        const e = new MouseEvent("mousemove", {
            view: window,
            bubbles: true,
            cancelable: true,
            clientX: x,
            clientY: y
        });

        document.dispatchEvent(e);
        log("log", `Dispatched mousemove event (X: ${x}, Y: ${y}).`);
    };

    createConsoleProxy();
    checkIfMechanismValid();
    setupIdleDetector();
    log("log", "Starting execution loop...");
    prepareExecute();
})();