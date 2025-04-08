// ==UserScript==
// @name         Bing Search Automator for Microsoft Rewards
// @namespace    http://tampermonkey.net/
// @version      1.9
// @description  Automate searches on Bing to earn Microsoft Rewards points
// @author       You
// @match        https://*.bing.com/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_addStyle
// @grant        GM_log
// @grant        GM_notification
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';
    
    // Initialize tracking variables
    let scriptInitialized = false;
    let buttonAdded = false;
    let stylesAdded = false;
    let initAttempts = 0;
    const MAX_INIT_ATTEMPTS = 5;
    
    // Setup custom console logging to distinguish our script's logs from Bing's errors
    const scriptPrefix = '[Rewards Automator] ';
    const logInfo = (message) => console.log(`%c${scriptPrefix}${message}`, 'color: #0078d7; font-weight: bold;');
    const logError = (message) => console.error(`%c${scriptPrefix}ERROR: ${message}`, 'color: #ff3333; font-weight: bold;');
    const logWarning = (message) => console.warn(`%c${scriptPrefix}WARNING: ${message}`, 'color: #ff9900; font-weight: bold;');
    const logDebug = (message) => isDebugMode() && console.log(`%c${scriptPrefix}DEBUG: ${message}`, 'color: #00aa00; font-weight: bold;');

    // Safely access GM functions
    function safeGM(fn, defaultValue, ...args) {
        try {
            if (typeof fn === 'function') {
                return fn(...args);
            }
            return defaultValue;
        } catch (e) {
            logError(`GM API error: ${e.message}`);
            return defaultValue;
        }
    }
    
    // Safely get/set values with fallback to localStorage if GM API fails
    function getValue(key, defaultValue) {
        try {
            if (typeof GM_getValue === 'function') {
                return GM_getValue(key, defaultValue);
            } else {
                const value = localStorage.getItem(`bing_rewards_${key}`);
                return value !== null ? JSON.parse(value) : defaultValue;
            }
        } catch (e) {
            logError(`Error accessing storage (getValue): ${e.message}`);
            return defaultValue;
        }
    }
    
    function setValue(key, value) {
        try {
            if (typeof GM_setValue === 'function') {
                GM_setValue(key, value);
            } else {
                localStorage.setItem(`bing_rewards_${key}`, JSON.stringify(value));
            }
        } catch (e) {
            logError(`Error accessing storage (setValue): ${e.message}`);
        }
    }
    
    // Debug mode toggle and check
    function isDebugMode() {
        return getValue('debugMode', false);
    }
    
    // Show browser notification
    function showNotification(title, text, timeout = 5000) {
        try {
            if (typeof GM_notification === 'function') {
                GM_notification({
                    title: title,
                    text: text,
                    timeout: timeout
                });
            } else {
                // Fallback to alert for critical messages or just log for info
                if (title.includes("Error")) {
                    alert(`${title}: ${text}`);
                } else {
                    logInfo(`${title}: ${text}`);
                }
            }
        } catch (e) {
            logError(`Error showing notification: ${e.message}`);
        }
    }

    // Log script initialization
    logInfo(`Initializing v1.9 - ${new Date().toLocaleString()}`);
    if (isDebugMode()) logDebug("Debug mode is ON");
    
    // Configuration
    const config = {
        minInterval: 15, // Minimum time between searches in seconds
        maxInterval: 40, // Maximum time between searches in seconds
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
    let isRunning = getValue('bingAutoSearchRunning', false);
    let searchTimer = null;
    let lastRunTime = getValue('lastRunTime', 0);
    let searchCount = getValue('searchCount', 0);
    const MAX_SEARCHES = 30; // Maximum number of searches to perform
    
    // Store button position
    let buttonPosition = getValue('buttonPosition', { x: 20, y: 20 });

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

    // Function to get a random keyword or generate one
    function getRandomKeyword() {
        return generateSearchQuery();
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
                // Get a random keyword
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
                setValue('searchCount', searchCount);
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
                setValue('lastRunTime', lastRunTime);
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
                setValue('bingAutoSearchRunning', true);
                searchCount = 0;
                setValue('searchCount', searchCount);
                updateFloatingButton();
                
                logInfo('Starting automated Bing searches');
                const initialInterval = getRandomInterval();
                logInfo(`First search in ${initialInterval/1000} seconds`);
                searchTimer = setTimeout(performSearch, initialInterval);
                lastRunTime = Date.now() + initialInterval;
                setValue('lastRunTime', lastRunTime);
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
                setValue('bingAutoSearchRunning', false);
                if (searchTimer) {
                    clearTimeout(searchTimer);
                }
                updateFloatingButton();
                logInfo('Stopped automated Bing searches');
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
                
                setValue('buttonPosition', buttonPosition);
                
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
                setValue('buttonPosition', buttonPosition);
                
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

    // Function to toggle debug mode
    function toggleDebugMode() {
        const newDebugMode = !isDebugMode();
        setValue('debugMode', newDebugMode);
        showNotification(
            "Debug Mode Changed", 
            `Debug mode ${newDebugMode ? 'enabled' : 'disabled'}. Refresh the page to apply.`
        );
    }
    
    // Function to reset all settings
    function resetAllSettings() {
        if (confirm("Are you sure you want to reset all settings?")) {
            // Clear all settings
            setValue('bingAutoSearchRunning', false);
            setValue('lastRunTime', 0);
            setValue('searchCount', 0);
            setValue('buttonPosition', { x: 20, y: 20 });
            // Force refresh the page
            location.reload();
        }
    }

    // Register menu commands for Tampermonkey/Violentmonkey
    function registerMenuCommands() {
        try {
            if (typeof GM_registerMenuCommand === 'function') {
                GM_registerMenuCommand('Start Auto-Search', startAutomation);
                GM_registerMenuCommand('Stop Auto-Search', stopAutomation);
                GM_registerMenuCommand('Toggle Debug Mode', toggleDebugMode);
                GM_registerMenuCommand('Reset All Settings', resetAllSettings);
                logDebug('Menu commands registered');
            } else {
                logWarning('GM_registerMenuCommand not available');
            }
        } catch (e) {
            logError(`Failed to register menu commands: ${e.message}`);
        }
    }

    // Add CSS to page using a more secure method 
    function addStyles() {
        if (stylesAdded) return true;
        
        try {
            logDebug('Adding styles');
            // Use GM_addStyle if available (safer and preferred method)
            if (typeof GM_addStyle === 'function') {
                GM_addStyle(`
                    #bing-auto-search-button {
                        position: fixed;
                        left: ${buttonPosition.x}px;
                        top: ${buttonPosition.y}px;
                        width: 80px;
                        height: 80px;
                        border-radius: 12px;
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
                        z-index: 99999;
                        user-select: none;
                        transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.3s ease;
                        border: ${isDebugMode() ? '2px solid yellow' : '2px solid rgba(255,255,255,0.3)'};
                        will-change: transform, left, top;
                        background: linear-gradient(135deg, #0078d7, #0063b1);
                        touch-action: none;
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
                        background: linear-gradient(135deg, #ff5555, #ff3333) !important;
                    }
                    #bing-auto-search-button[data-dragging="true"].searching {
                        animation: none;
                    }
                    #rewards-automator-debug {
                        position: fixed;
                        bottom: 10px;
                        right: 10px;
                        background: rgba(0,0,0,0.7);
                        color: lime;
                        padding: 5px 10px;
                        border-radius: 5px;
                        font-family: monospace;
                        z-index: 99999;
                        font-size: 12px;
                    }
                `);
                logDebug('Styles added with GM_addStyle');
            } else {
                // Fallback method: Create a text node
                const style = document.createElement('style');
                style.type = 'text/css';
                const css = `
                    #bing-auto-search-button {
                        position: fixed;
                        left: ${buttonPosition.x}px;
                        top: ${buttonPosition.y}px;
                        width: 80px;
                        height: 80px;
                        border-radius: 12px;
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
                        z-index: 99999;
                        user-select: none;
                        transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.3s ease;
                        border: ${isDebugMode() ? '2px solid yellow' : '2px solid rgba(255,255,255,0.3)'};
                        will-change: transform, left, top;
                        background: linear-gradient(135deg, #0078d7, #0063b1);
                        touch-action: none;
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
                        background: linear-gradient(135deg, #ff5555, #ff3333) !important;
                    }
                    #bing-auto-search-button[data-dragging="true"].searching {
                        animation: none;
                    }
                    #rewards-automator-debug {
                        position: fixed;
                        bottom: 10px;
                        right: 10px;
                        background: rgba(0,0,0,0.7);
                        color: lime;
                        padding: 5px 10px;
                        border-radius: 5px;
                        font-family: monospace;
                        z-index: 99999;
                        font-size: 12px;
                    }
                `;
                style.appendChild(document.createTextNode(css));
                
                // Try to add to head or body, whichever is available
                if (document.head) {
                    document.head.appendChild(style);
                } else if (document.body) {
                    document.body.appendChild(style);
                } else {
                    // If neither is available, wait for DOMContentLoaded
                    document.addEventListener('DOMContentLoaded', function() {
                        document.head.appendChild(style);
                    });
                    return false; // Indicate we need to try again later
                }
                logDebug('Styles added with fallback method');
            }
            
            stylesAdded = true;
            return true;
        } catch (e) {
            logError(`Error adding styles: ${e.message}`);
            return false;
        }
    }

    // Create a floating activation button
    function addFloatingButton() {
        if (buttonAdded) return true;
        
        try {
            // Check if button already exists
            if (document.getElementById('bing-auto-search-button')) {
                logDebug('Button already exists');
                buttonAdded = true;
                return true;
            }
            
            // Make sure we have a body to append to
            if (!document.body) {
                logDebug('Document body not available yet, will try again later');
                return false;
            }
            
            logDebug('Creating floating button');
            const floatingButton = document.createElement('div');
            floatingButton.id = 'bing-auto-search-button';
            
            // Create icon
            const iconSpan = document.createElement('span');
            iconSpan.style.cssText = `
                font-size: 20px;
                margin-bottom: 4px;
                pointer-events: none;
            `;
            iconSpan.innerHTML = isRunning ? '⏹️' : '▶️';
            
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
            
            // Apply class if running
            if (isRunning) {
                floatingButton.classList.add('searching');
            }
            
            buttonAdded = true;
            logInfo('Floating button added successfully');
            return true;
        } catch (err) {
            logError(`Error in addFloatingButton: ${err.message}`);
            return false;
        }
    }

    // Update floating button text and color
    function updateFloatingButton() {
        try {
            const button = document.getElementById('bing-auto-search-button');
            if (button) {
                // Update icon and text
                const iconSpan = button.querySelector('span:first-child');
                if (iconSpan) {
                    iconSpan.innerHTML = isRunning ? '⏹️' : '▶️';
                }
                
                const textSpan = button.querySelector('span:nth-child(2)');
                if (textSpan) {
                    textSpan.textContent = isRunning ? 'STOP' : 'START';
                }
                
                const counterSpan = document.getElementById('bing-search-counter');
                if (counterSpan) {
                    counterSpan.textContent = isRunning ? `${searchCount}/${MAX_SEARCHES}` : 'Rewards';
                }
                
                // Add/remove searching class for styling
                if (isRunning) {
                    button.classList.add('searching');
                } else {
                    button.classList.remove('searching');
                }
            } else if (buttonAdded) {
                // Button should exist but doesn't - recreate it
                buttonAdded = false;
                addFloatingButton();
            }
        } catch (err) {
            logError(`Error in updateFloatingButton: ${err.message}`);
        }
    }

    // Multiple initialization methods for better compatibility
    function initializeScript() {
        if (scriptInitialized) return;
        
        try {
            logInfo(`Initializing script on: ${window.location.href}`);
            initAttempts++;
            
            // Check if we're on Bing
            if (!window.location.hostname.includes('bing.com')) {
                logWarning('Not on Bing search page. Script will only work on Bing.com');
                return;
            }
            
            // Add styles first
            if (!stylesAdded) {
                if (!addStyles()) {
                    if (initAttempts < MAX_INIT_ATTEMPTS) {
                        logDebug(`Failed to add styles, will retry (attempt ${initAttempts}/${MAX_INIT_ATTEMPTS})`);
                        setTimeout(initializeScript, 200 * initAttempts); // Exponential backoff
                    }
                    return;
                }
            }
            
            // Then add button
            if (!buttonAdded) {
                if (!addFloatingButton()) {
                    if (initAttempts < MAX_INIT_ATTEMPTS) {
                        logDebug(`Failed to add button, will retry (attempt ${initAttempts}/${MAX_INIT_ATTEMPTS})`);
                        setTimeout(initializeScript, 200 * initAttempts); // Exponential backoff
                    }
                    return;
                }
            }
            
            // Register menu commands
            registerMenuCommands();
            
            // Check if we need to resume operation
            if (isRunning) {
                logInfo('Resuming previous search session');
                checkAndResumeOperation();
            }
            
            // Add debug indicator if in debug mode
            if (isDebugMode() && !document.getElementById('rewards-automator-debug')) {
                const debugElement = document.createElement('div');
                debugElement.id = 'rewards-automator-debug';
                debugElement.textContent = `Rewards Automator v1.9 Active | ${isRunning ? 'Running' : 'Idle'}`;
                document.body.appendChild(debugElement);
            }
            
            scriptInitialized = true;
            logInfo('Script initialization complete!');
            
            // Indicate successful setup with a notification in debug mode
            if (isDebugMode()) {
                showNotification('Rewards Automator Ready', 
                    `Script initialized successfully. ${isRunning ? 'Searches are active.' : 'Click START to begin searches.'}`);
            }
        } catch (err) {
            logError(`Failed to initialize script: ${err.message}`);
            if (initAttempts < MAX_INIT_ATTEMPTS) {
                setTimeout(initializeScript, 500); // Try again soon
            }
        }
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

    // Handle any unhandled exceptions
    window.addEventListener('error', function(event) {
        // Only log errors from our script, not Bing's errors
        if (event.filename && event.filename.includes('RewardsAutomator')) {
            logError(`Unhandled error: ${event.message} at ${event.filename}:${event.lineno}`);
        }
    });
    
    // Try to initialize as early as possible
    initializeScript();
    
    // Multiple event listeners for different script managers
    if (document.readyState === 'loading') {
        logDebug('Document still loading, adding DOMContentLoaded listener');
        document.addEventListener('DOMContentLoaded', initializeScript);
    }
    
    // Add load event as a backup
    window.addEventListener('load', function() {
        logDebug('Window load event fired');
        if (!scriptInitialized) {
            logInfo('Script not yet initialized, trying again on window load');
            initializeScript();
        }
    });
    
    // Handle visibility change (when tab becomes active/inactive)
    document.addEventListener('visibilitychange', function() {
        if (document.visibilityState === 'visible') {
            logInfo('Tab is now active');
            // Try to reinitialize if needed (handles cases where the script failed on first load)
            if (!scriptInitialized) {
                initializeScript();
            } else {
                updateFloatingButton();
                checkAndResumeOperation();
            }
        } else {
            logDebug('Tab is now inactive');
        }
    });
    
    // Try again after a bit in case the page loads slowly
    setTimeout(function() {
        if (!scriptInitialized) {
            logInfo('Script not initialized after initial timeout, trying final initialization');
            initializeScript();
        }
    }, 1000);
    
    // One final attempt after page should definitely be loaded
    setTimeout(function() {
        if (!scriptInitialized) {
            logWarning('Script STILL not initialized after long timeout. Making final attempt.');
            // Force retry with increased visibility for debugging
            initAttempts = 0; // Reset attempt counter
            initializeScript();
            
            // If it's still not working, show a visible error to the user
            if (!scriptInitialized) {
                showNotification(
                    'Rewards Automator Error', 
                    'Failed to initialize after multiple attempts. Try enabling debug mode or reset settings.',
                    10000
                );
            }
        }
    }, 5000);
})();
