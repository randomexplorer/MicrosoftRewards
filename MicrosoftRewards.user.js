// ==UserScript==
// @name         Bing Search Automator for Microsoft Rewards
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Automate searches on Bing to earn Microsoft Rewards points
// @author       You
// @match        https://*.bing.com/*
// @match        https://rewards.bing.com/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_addStyle
// @grant        GM_log
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // Setup custom console logging to distinguish our script's logs from Bing's errors
    const scriptPrefix = '[Rewards Automator] ';
    const logInfo = (message) => console.log(`%c${scriptPrefix}${message}`, 'color: #0078d7; font-weight: bold;');
    const logError = (message) => console.error(`%c${scriptPrefix}ERROR: ${message}`, 'color: #ff3333; font-weight: bold;');
    const logWarning = (message) => console.warn(`%c${scriptPrefix}WARNING: ${message}`, 'color: #ff9900; font-weight: bold;');

    // Log script initialization
    logInfo(`Initializing v1.8 - ${new Date().toLocaleString()}`);

    // Configuration
    const config = {
        minInterval: 15, // Minimum time between searches in seconds
        maxInterval: 40, // Maximum time between searches in seconds
        // Maximum search history to store (24 hours worth of searches)
        maxSearchHistorySize: 100,
        // Keywords components for generating natural search queries
        keywordComponents: {
            // Question prefixes
            questionPrefixes: [
                "how to", "how do I", "what is", "what are", "where can I find",
                "why does", "when will", "who invented", "which", "can I"
            ],

            // Search prefixes (not questions)
            searchPrefixes: [
                "best", "top", "affordable", "cheap", "popular", "trending",
                "new", "nearby", "how to make", "ways to"
            ],

            // Topic categories (nouns/subjects)
            topics: {
                technology: [
                    "smartphones", "laptops", "gaming PC", "smart TV", "headphones",
                    "wireless earbuds", "smartwatch", "tablet", "bluetooth speaker",
                    "mechanical keyboard", "webcam", "external hard drive"
                ],

                software: [
                    "Windows 11", "macOS", "Linux", "iOS 18", "Android", "Office 365",
                    "Photoshop", "Excel", "Word", "PowerPoint", "Visual Studio Code",
                    "Chrome", "Firefox", "Edge browser"
                ],

                programming: [
                    "JavaScript", "Python", "Java", "C++", "TypeScript", "React",
                    "Angular", "Node.js", "SQL", "HTML", "CSS", "Go language", "Rust"
                ],

                health: [
                    "workout", "diet", "nutrition", "vitamins", "protein", "yoga",
                    "meditation", "mental health", "sleep", "exercises", "running",
                    "weight training", "cardio"
                ],

                food: [
                    "recipes", "cooking", "restaurants", "baking", "meal prep",
                    "breakfast", "lunch", "dinner", "dessert", "coffee", "smoothies",
                    "pizza", "sushi", "Italian food"
                ],

                travel: [
                    "vacation", "flights", "hotels", "resorts", "beaches", "mountains",
                    "national parks", "Europe trip", "Asia tour", "road trip",
                    "cruises", "travel insurance", "passport"
                ],

                shopping: [
                    "online shopping", "Amazon deals", "discount codes", "sales",
                    "fashion", "shoes", "electronics", "furniture", "home decor",
                    "kitchen appliances", "clothing brands"
                ],

                entertainment: [
                    "movies", "TV shows", "streaming services", "Netflix", "Disney+",
                    "HBO Max", "music", "concerts", "books", "podcasts", "video games",
                    "board games", "theater"
                ]
            },

            // Modifiers/qualifiers
            modifiers: [
                "for beginners", "tutorial", "guide", "review", "comparison",
                "near me", "online", "2025", "reddit", "best of 2025",
                "worth it", "alternatives", "vs", "prices"
            ],

            // Seasonal and current topics (April 2025)
            seasonal: [
                "Earth Day 2025", "spring activities", "spring cleaning",
                "tax deadline 2025", "April events", "gardening tips spring",
                "spring fashion 2025", "spring break destinations"
            ],

            // Time-based searches
            timeBased: [
                "today", "this week", "this weekend", "this month", "April 2025",
                "upcoming", "schedule", "release date", "launch"
            ],

            // Common everyday searches
            everyday: [
                "weather forecast", "news today", "stock market", "traffic updates",
                "sports scores", "exchange rate", "calculator", "translate",
                "dictionary", "maps"
            ]
        }
    };

    // Variables to control execution
    let isRunning = GM_getValue('bingAutoSearchRunning', false);
    let searchTimer = null;
    let lastRunTime = GM_getValue('lastRunTime', 0);
    let searchCount = GM_getValue('searchCount', 0);
    const MAX_SEARCHES = 30; // Maximum number of searches to perform

    // Store button position
    let buttonPosition = GM_getValue('buttonPosition', { x: 20, y: 20 });

    // Search history tracking - load from storage or initialize empty
    let searchHistory = GM_getValue('searchHistory', []);

    // Tab coordination - to handle multiple tabs
    let activeTabId = GM_getValue('activeTabId', null);
    let thisTabId = Date.now().toString() + Math.floor(Math.random() * 10000).toString();
    let buttonVisibleTabId = GM_getValue('buttonVisibleTabId', null);
    let lastButtonVisibilityUpdateTime = GM_getValue('lastButtonVisibilityUpdateTime', 0);
    const TAB_TIMEOUT_MS = 30000; // Reduced to 30 seconds to detect stale tabs faster
    const HEARTBEAT_INTERVAL = 10000; // Heartbeat every 10 seconds

    // Function to check if this tab should be active
    function isActiveTab() {
        return activeTabId === thisTabId || activeTabId === null;
    }

    // Function to check if this tab should show the button
    function shouldShowButtonInThisTab() {
        // Reset button visibility if it's been too long since the last update
        const currentTime = Date.now();
        if (buttonVisibleTabId !== null &&
            (currentTime - lastButtonVisibilityUpdateTime) > TAB_TIMEOUT_MS) {
            logInfo('Button visibility timeout exceeded, resetting...');
            GM_setValue('buttonVisibleTabId', null);
            buttonVisibleTabId = null;
        }

        // If no tab is currently showing the button, this tab can show it
        if (buttonVisibleTabId === null) {
            return true;
        }

        // If this tab is already designated to show the button, continue showing it
        if (buttonVisibleTabId === thisTabId) {
            return true;
        }

        // If another tab is showing the button, don't show it here
        return false;
    }

    // Function to claim the right to show the button in this tab
    function claimButtonVisibility() {
        buttonVisibleTabId = thisTabId;
        GM_setValue('buttonVisibleTabId', thisTabId);
        lastButtonVisibilityUpdateTime = Date.now();
        GM_setValue('lastButtonVisibilityUpdateTime', lastButtonVisibilityUpdateTime);
        logInfo('This tab claimed button visibility');
    }

    // Function to update last button visibility time
    function updateButtonVisibilityTime() {
        if (buttonVisibleTabId === thisTabId) {
            lastButtonVisibilityUpdateTime = Date.now();
            GM_setValue('lastButtonVisibilityUpdateTime', lastButtonVisibilityUpdateTime);
        }
    }

    // Function to claim active status for this tab
    function claimActiveStatus() {
        activeTabId = thisTabId;
        GM_setValue('activeTabId', thisTabId);
    }

    // Function to release active status
    function releaseActiveStatus() {
        if (isActiveTab()) {
            GM_setValue('activeTabId', null);
        }
    }

    // Function to release button visibility
    function releaseButtonVisibility() {
        if (buttonVisibleTabId === thisTabId) {
            GM_setValue('buttonVisibleTabId', null);
        }
    }

    // Function to reset tab coordination if needed
    function resetTabCoordination() {
        GM_setValue('buttonVisibleTabId', null);
        GM_setValue('activeTabId', null);
        buttonVisibleTabId = null;
        activeTabId = null;
        logInfo('Tab coordination reset');
    }

    // Register keyboard shortcut for resetting coordination
    document.addEventListener('keydown', function(e) {
        // Alt+Shift+R to reset tab coordination
        if (e.altKey && e.shiftKey && e.key === 'R') {
            resetTabCoordination();
            setTimeout(() => {
                claimButtonVisibility();
                addFloatingButton(); // Re-add the button after reset
            }, 100);
            logInfo('Tab coordination reset triggered by keyboard shortcut');
        }
    });

    // Check tab status periodically to handle closed tabs
    function checkTabStatus() {
        try {
            const currentTime = Date.now();

            // Check if button visibility claim is stale
            if (buttonVisibleTabId !== null &&
                buttonVisibleTabId !== thisTabId &&
                (currentTime - lastButtonVisibilityUpdateTime) > TAB_TIMEOUT_MS) {
                logInfo('Detected stale button visibility claim, resetting...');
                GM_setValue('buttonVisibleTabId', null);
                buttonVisibleTabId = null;
            }

            // If automation is not running anywhere, reset the active tab
            if (!isRunning && activeTabId !== null) {
                releaseActiveStatus();
            }

            // If automation is running but no tab is active, claim it
            if (isRunning && activeTabId === null) {
                claimActiveStatus();
            }

            // If no tab is showing the button, claim visibility
            if (buttonVisibleTabId === null) {
                claimButtonVisibility();
                // Force redraw of button
                const existingButton = document.getElementById('bing-auto-search-button');
                if (!existingButton) {
                    addFloatingButton();
                }
            } else if (buttonVisibleTabId === thisTabId) {
                // If this tab has the button, update the timestamp
                updateButtonVisibilityTime();
            }

            // Schedule next check
            setTimeout(checkTabStatus, 5000); // Check every 5 seconds
        } catch (err) {
            logError(`Error in checkTabStatus: ${err.message}`);
        }
    }

    // Document visibility change handler - update button visibility timestamp when tab is active
    document.addEventListener('visibilitychange', function() {
        if (document.visibilityState === 'visible') {
            if (buttonVisibleTabId === thisTabId) {
                updateButtonVisibilityTime();
            } else if (buttonVisibleTabId === null) {
                claimButtonVisibility();
                const existingButton = document.getElementById('bing-auto-search-button');
                if (!existingButton) {
                    addFloatingButton();
                }
            } else {
                // Check if the button visibility claim is stale
                const currentTime = Date.now();
                if ((currentTime - lastButtonVisibilityUpdateTime) > TAB_TIMEOUT_MS) {
                    logInfo('Detected stale button visibility on tab activation, claiming...');
                    GM_setValue('buttonVisibleTabId', null);
                    buttonVisibleTabId = null;
                    setTimeout(() => {
                        claimButtonVisibility();
                        addFloatingButton();
                    }, 100);
                }
            }
        }
    });

    // Handle tab close/unload
    window.addEventListener('beforeunload', function() {
        // If this tab is showing the button, release it so other tabs can show it
        if (buttonVisibleTabId === thisTabId) {
            releaseButtonVisibility();
        }

        // If this tab is running the automation, stop it
        if (isActiveTab() && isRunning) {
            releaseActiveStatus();
        }
    });

    // Function to check if a search query has been used recently
    function isQueryInHistory(query) {
        return searchHistory.includes(query.toLowerCase());
    }

    // Function to add a query to search history
    function addQueryToHistory(query) {
        // Convert to lowercase for case-insensitive comparison
        query = query.toLowerCase();

        // Add to history if not already present
        if (!isQueryInHistory(query)) {
            searchHistory.push(query);

            // Trim history to max size
            if (searchHistory.length > config.maxSearchHistorySize) {
                searchHistory = searchHistory.slice(-config.maxSearchHistorySize);
            }

            // Save to persistent storage
            GM_setValue('searchHistory', searchHistory);
        }
    }

    // Function to clear search history
    function clearSearchHistory() {
        searchHistory = [];
        GM_setValue('searchHistory', searchHistory);
        logInfo('Search history cleared');
    }

    // Helper function to get random item from array
    function getRandomItem(array) {
        return array[Math.floor(Math.random() * array.length)];
    }

    // Function to pick random topic
    function getRandomTopic() {
        try {
            const categories = Object.keys(config.keywordComponents.topics);
            const randomCategory = getRandomItem(categories);
            return getRandomItem(config.keywordComponents.topics[randomCategory]);
        } catch (e) {
            logError(`Error in getRandomTopic: ${e.message}`);
            return "error getting topic";
        }
    }

    // Function to generate a natural search query using components
    function generateSearchQuery() {
        // Choose query pattern based on weighted random selection
        const rand = Math.random();

        // 25% chance: Question pattern
        if (rand < 0.25) {
            const prefix = getRandomItem(config.keywordComponents.questionPrefixes);
            const topic = getRandomTopic();

            // Sometimes add a modifier
            if (Math.random() < 0.3) {
                const modifier = getRandomItem(config.keywordComponents.modifiers);
                return `${prefix} ${topic} ${modifier}`;
            } else {
                return `${prefix} ${topic}`;
            }
        }
        // 20% chance: Search prefix + topic + modifier
        else if (rand < 0.45) {
            const prefix = getRandomItem(config.keywordComponents.searchPrefixes);
            const topic = getRandomTopic();
            const modifier = getRandomItem(config.keywordComponents.modifiers);
            return `${prefix} ${topic} ${modifier}`;
        }
        // 20% chance: Search prefix + topic
        else if (rand < 0.65) {
            const prefix = getRandomItem(config.keywordComponents.searchPrefixes);
            const topic = getRandomTopic();
            return `${prefix} ${topic}`;
        }
        // 10% chance: Topic + modifier
        else if (rand < 0.75) {
            const topic = getRandomTopic();
            const modifier = getRandomItem(config.keywordComponents.modifiers);
            return `${topic} ${modifier}`;
        }
        // 10% chance: Seasonal/current topic
        else if (rand < 0.85) {
            return getRandomItem(config.keywordComponents.seasonal);
        }
        // 5% chance: Time-based search
        else if (rand < 0.90) {
            const topic = getRandomTopic();
            const timeModifier = getRandomItem(config.keywordComponents.timeBased);
            return `${topic} ${timeModifier}`;
        }
        // 10% chance: Everyday search
        else {
            return getRandomItem(config.keywordComponents.everyday);
        }
    }

    // Function to get a unique random keyword that hasn't been used recently
    function getRandomKeyword() {
        // Set a limit to prevent infinite loops
        const MAX_ATTEMPTS = 10;
        let attempts = 0;
        let keyword;

        do {
            keyword = generateSearchQuery();
            attempts++;

            // If we've tried too many times, just use this keyword and make it slightly unique
            if (attempts >= MAX_ATTEMPTS && isQueryInHistory(keyword)) {
                const uniqueSuffix = Math.floor(Math.random() * 100);
                keyword = `${keyword} ${uniqueSuffix}`;
                break;
            }
        } while (isQueryInHistory(keyword) && attempts < MAX_ATTEMPTS);

        // Add to history for tracking
        addQueryToHistory(keyword);

        return keyword;
    }

    // Function to get a random interval between minInterval and maxInterval
    function getRandomInterval() {
        return Math.floor(Math.random() * (config.maxInterval - config.minInterval + 1) + config.minInterval) * 1000;
    }

    // Function to check if we are on Bing search page
    function isOnBingSearchPage() {
        return window.location.hostname.includes('bing.com');
    }

    // Function to perform a search
    function performSearch() {
        try {
            if (!isOnBingSearchPage()) {
                logInfo('Not on Bing search page, opening Bing...');
                window.open('https://www.bing.com', '_blank');
                return;
            }

            // Find the search input field
            const searchInput = document.querySelector('input[name="q"]') ||
                                document.querySelector('#sb_form_q') ||
                                document.querySelector('input[type="search"]');

            if (searchInput) {
                // Get a random keyword that hasn't been used recently
                const keyword = getRandomKeyword();

                // Set the value and trigger events
                searchInput.value = keyword;
                searchInput.dispatchEvent(new Event('input', { bubbles: true }));

                // Find and click the search button
                const searchButton = document.querySelector('#search_icon') ||
                                    document.querySelector('button[type="submit"]') ||
                                    document.querySelector('#sb_form_go');

                if (searchButton) {
                    searchButton.click();
                } else {
                    // If no button is found, try to submit the form
                    const form = searchInput.closest('form');
                    if (form) {
                        form.submit();
                    }
                }

                searchCount++;
                GM_setValue('searchCount', searchCount);
                updateFloatingButton();
                logInfo(`Performed search ${searchCount}/${MAX_SEARCHES} for: ${keyword}`);
            } else {
                logError('Search input not found');
            }

            // Schedule the next search if still running
            if (isRunning && searchCount < MAX_SEARCHES) {
                const nextInterval = getRandomInterval();
                logInfo(`Next search in ${nextInterval/1000} seconds`);
                searchTimer = setTimeout(performSearch, nextInterval);
                lastRunTime = Date.now() + nextInterval;
                GM_setValue('lastRunTime', lastRunTime);
            } else if (searchCount >= MAX_SEARCHES) {
                logInfo('Reached maximum number of searches. Stopping automation.');
                stopAutomation();
            }
        } catch (err) {
            logError(`Error in performSearch: ${err.message}`);
        }
    }

    // Function to start automated searches
    function startAutomation() {
        try {
            if (!isRunning) {
                if (!isOnBingSearchPage()) {
                    alert('Please navigate to Bing search engine first, then activate the script.');
                    window.open('https://www.bing.com', '_blank');
                    return;
                }

                isRunning = true;
                GM_setValue('bingAutoSearchRunning', true);
                searchCount = 0;
                GM_setValue('searchCount', searchCount);
                updateFloatingButton();

                logInfo('Starting automated Bing searches');
                const initialInterval = getRandomInterval();
                logInfo(`First search in ${initialInterval/1000} seconds`);
                searchTimer = setTimeout(performSearch, initialInterval);
                lastRunTime = Date.now() + initialInterval;
                GM_setValue('lastRunTime', lastRunTime);

                // Claim active status for this tab
                claimActiveStatus();
            }
        } catch (err) {
            logError(`Error in startAutomation: ${err.message}`);
        }
    }

    // Function to stop automated searches
    function stopAutomation() {
        try {
            if (isRunning) {
                isRunning = false;
                GM_setValue('bingAutoSearchRunning', false);
                if (searchTimer) {
                    clearTimeout(searchTimer);
                }
                updateFloatingButton();
                logInfo('Stopped automated Bing searches');

                // Release active status for this tab
                releaseActiveStatus();
            }
        } catch (err) {
            logError(`Error in stopAutomation: ${err.message}`);
        }
    }

    // Make an element draggable with smooth movement
    function makeDraggable(element) {
        // Global variables for tracking drag state
        let isDragging = false;
        let dragStartX, dragStartY;
        let elementStartLeft, elementStartTop;
        let lastFrameTime = 0;

        // Get the computed style to handle transitions properly
        const computedStyle = window.getComputedStyle(element);

        // Setup direct positioning for smoother dragging
        function setupForDragging() {
            // Store original values
            const originalTransition = element.style.transition;

            // Temporarily remove transitions for smooth drag
            element.style.transition = 'none';

            // Force a reflow to apply the style change immediately
            void element.offsetWidth;

            return function() {
                // Restore original transition
                element.style.transition = originalTransition;
            };
        }

        // Handler for mouse down event to start dragging
        element.addEventListener('mousedown', function(e) {
            // Prevent text selection during drag
            e.preventDefault();

            // Check if we're targeting a child element - if so, only handle drag if it has pointer-events: none
            if (e.target !== element) {
                const childStyle = window.getComputedStyle(e.target);
                if (childStyle.pointerEvents !== 'none') {
                    // Let the click go through to the child
                    return;
                }
            }

            // Setup for smooth dragging
            const restoreStyles = setupForDragging();

            // Get current position and cursor location
            const rect = element.getBoundingClientRect();
            dragStartX = e.clientX;
            dragStartY = e.clientY;

            // Store the exact starting position (using exact pixel values)
            elementStartLeft = rect.left;
            elementStartTop = rect.top;

            // Set the element position precisely
            element.style.left = elementStartLeft + 'px';
            element.style.top = elementStartTop + 'px';

            // Mark as dragging
            isDragging = true;
            element.style.opacity = '0.8';
            element.style.cursor = 'grabbing';

            // Listen for mouse up to stop dragging
            const mouseUpHandler = function(e) {
                if (!isDragging) return;

                isDragging = false;
                element.style.opacity = '1';
                element.style.cursor = 'pointer';

                // Save final position
                buttonPosition = {
                    x: parseInt(element.style.left) || buttonPosition.x,
                    y: parseInt(element.style.top) || buttonPosition.y
                };

                GM_setValue('buttonPosition', buttonPosition);

                // Clean up
                restoreStyles();
                document.removeEventListener('mouseup', mouseUpHandler);
                document.removeEventListener('mousemove', mouseMoveHandler);
            };

            // Use requestAnimationFrame for smooth motion
            const mouseMoveHandler = function(e) {
                if (!isDragging) return;

                // Use requestAnimationFrame for smoother performance
                requestAnimationFrame(function() {
                    // Calculate new position with precise offset from starting point
                    const deltaX = e.clientX - dragStartX;
                    const deltaY = e.clientY - dragStartY;

                    // Apply the exact movement distance
                    let newLeft = elementStartLeft + deltaX;
                    let newTop = elementStartTop + deltaY;

                    // Keep within viewport bounds
                    const maxX = window.innerWidth - element.offsetWidth;
                    const maxY = window.innerHeight - element.offsetHeight;

                    newLeft = Math.min(Math.max(0, newLeft), maxX);
                    newTop = Math.min(Math.max(0, newTop), maxY);

                    // Apply new position directly for maximum smoothness
                    element.style.left = newLeft + 'px';
                    element.style.top = newTop + 'px';
                });
            };

            document.addEventListener('mouseup', mouseUpHandler);
            document.addEventListener('mousemove', mouseMoveHandler);
        });

        // Touch event handlers with the same smooth animation approach
        element.addEventListener('touchstart', function(e) {
            // Prevent scrolling while dragging
            e.preventDefault();

            // Setup for smooth dragging
            const restoreStyles = setupForDragging();

            // Get touch and position information
            const touch = e.touches[0];
            const rect = element.getBoundingClientRect();

            dragStartX = touch.clientX;
            dragStartY = touch.clientY;
            elementStartLeft = rect.left;
            elementStartTop = rect.top;

            // Set the element position precisely
            element.style.left = elementStartLeft + 'px';
            element.style.top = elementStartTop + 'px';

            isDragging = true;
            element.style.opacity = '0.8';

            const touchEndHandler = function() {
                if (!isDragging) return;

                isDragging = false;
                element.style.opacity = '1';

                // Save final position
                buttonPosition = {
                    x: parseInt(element.style.left) || buttonPosition.x,
                    y: parseInt(element.style.top) || buttonPosition.y
                };
                GM_setValue('buttonPosition', buttonPosition);

                // Clean up
                restoreStyles();
                document.removeEventListener('touchend', touchEndHandler);
                document.removeEventListener('touchcancel', touchEndHandler);
                document.removeEventListener('touchmove', touchMoveHandler);
            };

            const touchMoveHandler = function(e) {
                if (!isDragging) return;
                e.preventDefault(); // Prevent scrolling

                requestAnimationFrame(function() {
                    const touch = e.touches[0];

                    // Calculate new position with precise offset from starting point
                    const deltaX = touch.clientX - dragStartX;
                    const deltaY = touch.clientY - dragStartY;

                    // Apply the exact movement distance
                    let newLeft = elementStartLeft + deltaX;
                    let newTop = elementStartTop + deltaY;

                    // Keep within viewport bounds
                    const maxX = window.innerWidth - element.offsetWidth;
                    const maxY = window.innerHeight - element.offsetHeight;

                    newLeft = Math.min(Math.max(0, newLeft), maxX);
                    newTop = Math.min(Math.max(0, newTop), maxY);

                    // Apply new position directly for maximum smoothness
                    element.style.left = newLeft + 'px';
                    element.style.top = newTop + 'px';
                });
            };

            document.addEventListener('touchend', touchEndHandler);
            document.addEventListener('touchcancel', touchEndHandler);
            document.addEventListener('touchmove', touchMoveHandler);
        });
    }

    // Function to handle button click (separate from dragging)
    function handleButtonClick(element) {
        let startX, startY, hasMoved = false;
        const clickThreshold = 5; // pixels

        element.addEventListener('mousedown', function(e) {
            startX = e.clientX;
            startY = e.clientY;
            hasMoved = false;
        });

        element.addEventListener('mousemove', function(e) {
            if (!startX) return;

            const deltaX = Math.abs(e.clientX - startX);
            const deltaY = Math.abs(e.clientY - startY);

            if (deltaX > clickThreshold || deltaY > clickThreshold) {
                hasMoved = true;
            }
        });

        element.addEventListener('mouseup', function(e) {
            if (!hasMoved) {
                if (isRunning) {
                    stopAutomation();
                } else {
                    startAutomation();
                }
            }

            startX = startY = null;
            hasMoved = false;
        });

        // Touch version of the same logic
        element.addEventListener('touchstart', function(e) {
            startX = e.touches[0].clientX;
            startY = e.touches[0].clientY;
            hasMoved = false;
        });

        element.addEventListener('touchmove', function(e) {
            if (!startX) return;

            const deltaX = Math.abs(e.touches[0].clientX - startX);
            const deltaY = Math.abs(e.touches[0].clientY - startY);

            if (deltaX > clickThreshold || deltaY > clickThreshold) {
                hasMoved = true;
            }
        });

        element.addEventListener('touchend', function(e) {
            if (!hasMoved) {
                if (isRunning) {
                    stopAutomation();
                } else {
                    startAutomation();
                }
            }

            startX = startY = null;
            hasMoved = false;
        });
    }

    // Create a floating activation button
    function addFloatingButton() {
        try {
            // Check if button already exists to prevent duplicates
            if (document.getElementById('bing-auto-search-button')) {
                logInfo('Button already exists, not creating a duplicate');
                return document.getElementById('bing-auto-search-button');
            }

            // Check if this tab should show the button
            if (!shouldShowButtonInThisTab()) {
                logInfo('This tab is not designated to show the button');
                return null;
            }

            const floatingButton = document.createElement('div');
            floatingButton.id = 'bing-auto-search-button';

            // Set position using saved coordinates
            floatingButton.style.cssText = `
                position: fixed;
                left: ${buttonPosition.x}px;
                top: ${buttonPosition.y}px;
                width: 80px;
                height: 80px;
                border-radius: 12px;
                background: ${isRunning ?
                    'linear-gradient(135deg, #ff5555, #ff3333)' :
                    'linear-gradient(135deg, #0078d7, #0063b1)'};
                color: white;
                font-size: 14px;
                font-family: Arial, sans-serif;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                cursor: pointer;
                box-shadow: 0 4px 10px rgba(0,0,0,0.3);
                z-index: 9999;
                user-select: none;
                transition: all 0.3s ease;
                border: 2px solid rgba(255,255,255,0.3);
                will-change: transform, left, top;
            `;

            // Create icon - Fix for TrustedHTML error
            const iconSpan = document.createElement('span');
            iconSpan.style.cssText = `
                font-size: 20px;
                margin-bottom: 4px;
                pointer-events: none;
            `;
            // Use textContent instead of innerHTML
            iconSpan.textContent = isRunning ? '⏹️' : '▶️';

            // Create text
            const textSpan = document.createElement('span');
            textSpan.style.cssText = `
                font-weight: bold;
                pointer-events: none;
            `;
            textSpan.textContent = isRunning ? 'STOP' : 'START';

            // Create counter
            const counterSpan = document.createElement('span');
            counterSpan.id = 'bing-search-counter';
            counterSpan.style.cssText = `
                font-size: 12px;
                margin-top: 2px;
                opacity: 0.9;
                pointer-events: none;
            `;
            counterSpan.textContent = isRunning ? `${searchCount}/${MAX_SEARCHES}` : 'Rewards';

            floatingButton.appendChild(iconSpan);
            floatingButton.appendChild(textSpan);
            floatingButton.appendChild(counterSpan);

            // Add hover effect
            floatingButton.addEventListener('mouseenter', function() {
                // Only apply hover effect if not dragging
                if (!document.querySelector('#bing-auto-search-button[data-dragging="true"]')) {
                    this.style.transform = 'scale(1.05)';
                    this.style.boxShadow = '0 6px 15px rgba(0,0,0,0.4)';
                }
            });

            floatingButton.addEventListener('mouseleave', function() {
                this.style.transform = 'scale(1)';
                this.style.boxShadow = '0 4px 10px rgba(0,0,0,0.3)';
            });

            document.body.appendChild(floatingButton);

            // Make button draggable
            makeDraggable(floatingButton);

            // Handle button click separately from dragging
            handleButtonClick(floatingButton);

            // Claim button visibility for this tab
            claimButtonVisibility();

            logInfo('Floating button added successfully');
            return floatingButton;
        } catch (err) {
            logError(`Error in addFloatingButton: ${err.message}`);
            return null;
        }
    }

    // Update floating button text and color
    function updateFloatingButton() {
        try {
            const button = document.getElementById('bing-auto-search-button');
            if (button) {
                // Update background gradient
                button.style.background = isRunning ?
                    'linear-gradient(135deg, #ff5555, #ff3333)' :
                    'linear-gradient(135deg, #0078d7, #0063b1)';

                // Update icon and text - Fix for TrustedHTML error
                const iconSpan = button.querySelector('span:first-child');
                if (iconSpan) {
                    // Use textContent instead of innerHTML
                    iconSpan.textContent = isRunning ? '⏹️' : '▶️';
                }

                const textSpan = button.querySelector('span:nth-child(2)');
                if (textSpan) {
                    textSpan.textContent = isRunning ? 'STOP' : 'START';
                }

                const counterSpan = document.getElementById('bing-search-counter');
                if (counterSpan) {
                    counterSpan.textContent = isRunning ? `${searchCount}/${MAX_SEARCHES}` : 'Rewards';
                }

                // Add/remove pulsing effect
                if (isRunning) {
                    button.classList.add('searching');
                } else {
                    button.classList.remove('searching');
                }
            }
        } catch (err) {
            logError(`Error in updateFloatingButton: ${err.message}`);
        }
    }

    // Register menu commands for Tampermonkey
    function registerMenuCommands() {
        GM_registerMenuCommand('Start Auto-Search', startAutomation);
        GM_registerMenuCommand('Stop Auto-Search', stopAutomation);
        GM_registerMenuCommand('Clear Search History', clearSearchHistory);
    }

    // Check if we need to resume operation (when tab becomes active again)
    function checkAndResumeOperation() {
        try {
            if (isRunning && !searchTimer) {
                const currentTime = Date.now();
                if (currentTime >= lastRunTime) {
                    // The scheduled time has passed, perform search immediately
                    performSearch();
                } else {
                    // Schedule the remaining time
                    const remainingTime = lastRunTime - currentTime;
                    logInfo(`Resuming operation, next search in ${remainingTime/1000} seconds`);
                    searchTimer = setTimeout(performSearch, remainingTime);
                }
            }
        } catch (err) {
            logError(`Error in checkAndResumeOperation: ${err.message}`);
        }
    }

    // Add CSS to page using a more secure method that works with TrustedHTML restrictions
    function addStyles() {
        try {
            logInfo('Adding styles');
            // Use GM_addStyle if available (safer and preferred method)
            if (typeof GM_addStyle !== 'undefined') {
                GM_addStyle(`
                    #bing-auto-search-button {
                        transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.3s ease;
                    }
                    #bing-auto-search-button:active {
                        transform: scale(0.95);
                    }
                    @keyframes pulse {
                        0% { transform: scale(1); }
                        50% { transform: scale(1.05); }
                        100% { transform: scale(1); }
                    }
                    .searching {
                        animation: pulse 2s infinite;
                    }
                    #bing-auto-search-button[data-dragging="true"].searching {
                        animation: none;
                    }
                    #bing-auto-search-button {
                        touch-action: none;
                    }
                `);
            } else {
                // Fallback method 1: Create a text node instead of using innerHTML
                const style = document.createElement('style');
                style.type = 'text/css';
                const css = `
                    #bing-auto-search-button {
                        transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.3s ease;
                    }
                    #bing-auto-search-button:active {
                        transform: scale(0.95);
                    }
                    @keyframes pulse {
                        0% { transform: scale(1); }
                        50% { transform: scale(1.05); }
                        100% { transform: scale(1); }
                    }
                    .searching {
                        animation: pulse 2s infinite;
                    }
                    #bing-auto-search-button[data-dragging="true"].searching {
                        animation: none;
                    }
                    #bing-auto-search-button {
                        touch-action: none;
                    }
                `;
                // Create a text node instead of using innerHTML
                style.appendChild(document.createTextNode(css));
                document.head.appendChild(style);
            }
            logInfo('Styles added successfully');
        } catch (e) {
            logError(`Error adding styles: ${e.message}`);
            // Final fallback: Apply styles directly to the element when created
        }
    }

    // Handle any unhandled exceptions
    window.addEventListener('error', function(event) {
        // Only log errors from our script, not Bing's errors
        if (event.filename && event.filename.includes('microsoftRewards.js')) {
            logError(`Unhandled error: ${event.message} at ${event.filename}:${event.lineno}`);
        }
    });

    // Set up heartbeat to keep tab status updated
    function startHeartbeat() {
        // Immediate first heartbeat
        sendHeartbeat();

        // Schedule regular heartbeats
        setInterval(sendHeartbeat, HEARTBEAT_INTERVAL);
    }

    // Send heartbeat to indicate this tab is active
    function sendHeartbeat() {
        if (document.visibilityState === 'visible') {
            // Update last activity time if this tab has the button
            if (buttonVisibleTabId === thisTabId) {
                updateButtonVisibilityTime();
                logInfo('Heartbeat sent - this tab has the button');
            }

            // If no tab claims to have the button but we're visible, try to claim it
            if (buttonVisibleTabId === null) {
                claimButtonVisibility();
                const existingButton = document.getElementById('bing-auto-search-button');
                if (!existingButton) {
                    addFloatingButton();
                }
            }
        }
    }

    // Function to observe DOM changes and manage button state
    function observeBodyChanges() {
        try {
            const observer = new MutationObserver(mutations => {
                // Check if our button still exists
                const buttonExists = !!document.getElementById('bing-auto-search-button');

                // If button doesn't exist, recreate it
                if (!buttonExists && shouldShowButtonInThisTab()) {
                    logInfo('Button disappeared, recreating...');
                    addFloatingButton();
                }
            });

            // Start observing the document body for DOM changes
            observer.observe(document.body, {
                childList: true,
                subtree: true
            });

            logInfo('DOM observer started');
        } catch (err) {
            logError(`Error in observeBodyChanges: ${err.message}`);
        }
    }

    // Ensure button stays within viewport bounds
    function ensureButtonInViewport() {
        try {
            const button = document.getElementById('bing-auto-search-button');
            if (button) {
                const rect = button.getBoundingClientRect();
                const maxX = window.innerWidth - button.offsetWidth;
                const maxY = window.innerHeight - button.offsetHeight;
                
                // If button is outside viewport, reposition it
                let needsRepositioning = false;
                let newLeft = parseInt(button.style.left) || buttonPosition.x;
                let newTop = parseInt(button.style.top) || buttonPosition.y;
                
                if (newLeft > maxX) {
                    newLeft = Math.max(0, maxX);
                    needsRepositioning = true;
                }
                
                if (newTop > maxY) {
                    newTop = Math.max(0, maxY);
                    needsRepositioning = true;
                }
                
                if (needsRepositioning) {
                    logInfo(`Repositioning button from (${button.style.left}, ${button.style.top}) to (${newLeft}px, ${newTop}px)`);
                    button.style.left = `${newLeft}px`;
                    button.style.top = `${newTop}px`;
                    
                    // Update stored position
                    buttonPosition.x = newLeft;
                    buttonPosition.y = newTop;
                    GM_setValue('buttonPosition', buttonPosition);
                }
            }
        } catch (err) {
            logError(`Error in ensureButtonInViewport: ${err.message}`);
        }
    }

    // Handle window resize events
    function handleWindowResize() {
        ensureButtonInViewport();
    }

    // Handle navigation events (URL changes)
    function handleNavigation() {
        try {
            // Check if button exists and is visible
            const buttonExists = !!document.getElementById('bing-auto-search-button');
            
            if (!buttonExists && shouldShowButtonInThisTab()) {
                logInfo('Navigation detected, ensuring button is visible...');
                addFloatingButton();
            }
            
            // Ensure button is within viewport even if it exists
            ensureButtonInViewport();
        } catch (err) {
            logError(`Error in handleNavigation: ${err.message}`);
        }
    }
    
    // Monitor for URL changes
    let lastUrl = window.location.href;
    function checkUrlChange() {
        const currentUrl = window.location.href;
        if (currentUrl !== lastUrl) {
            lastUrl = currentUrl;
            handleNavigation();
        }
    }

    // Initialize when the page is fully loaded
    window.addEventListener('load', function() {
        logInfo('Bing Search Automator initialized');
        addStyles();

        // Use a short delay before adding the button to ensure DOM is stable
        setTimeout(() => {
            // Try to add the button
            const button = addFloatingButton();

            // If button wasn't added and no tab is showing it, force reset
            if (!button && !buttonVisibleTabId) {
                resetTabCoordination();
                // Try again after reset
                setTimeout(() => {
                    addFloatingButton();
                }, 200);
            }

            // Ensure button is within viewport initially
            ensureButtonInViewport();

            // Add window resize event listener
            window.addEventListener('resize', handleWindowResize);
            
            // Set up URL change monitoring
            setInterval(checkUrlChange, 1000);

            // Check and handle dynamic content loading (like when clicking trophy icon)
            observeBodyChanges();

            // Check if we need to resume operation
            if (isRunning) {
                checkAndResumeOperation();
            }

            // Start heartbeat to keep tab status updated
            startHeartbeat();

            // Check and update tab status
            checkTabStatus();
        }, 500);

        registerMenuCommands();

        // Add manual reset command
        GM_registerMenuCommand('Reset Tab Coordination', resetTabCoordination);

        // Add message about console errors not being related to our script
        logInfo('NOTE: Console errors from Bing\'s own scripts are normal and not related to this automation tool');

        // Log search history stats
        logInfo(`Search history contains ${searchHistory.length} entries`);
    });
})();
