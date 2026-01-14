/**
 * SharePoint Site Collection URL Extractor
 * Efficiently extracts site collection URLs from SharePoint URLs
 */

(function() {
    'use strict';

    // ============================================
    // Configuration
    // ============================================
    const CONFIG = {
        DEBOUNCE_DELAY: 500,      // ms to wait after typing stops
        BATCH_SIZE: 1000,          // URLs to process per batch
        TOAST_DURATION: 2000,      // ms to show toast notification
        PROGRESS_THRESHOLD: 100    // Show progress bar if more than this many URLs
    };

    // ============================================
    // DOM Elements
    // ============================================
    const elements = {
        urlInput: document.getElementById('urlInput'),
        inputCount: document.getElementById('inputCount'),
        clearBtn: document.getElementById('clearBtn'),
        progressBar: document.getElementById('progressBar'),
        progressFill: document.getElementById('progressFill'),
        progressText: document.getElementById('progressText'),
        convertedOutput: document.getElementById('convertedOutput'),
        convertedCount: document.getElementById('convertedCount'),
        copyConverted: document.getElementById('copyConverted'),
        uniqueList: document.getElementById('uniqueList'),
        uniqueCount: document.getElementById('uniqueCount'),
        copyUnique: document.getElementById('copyUnique'),
        toast: document.getElementById('toast'),
        toastMessage: document.getElementById('toastMessage')
    };

    // ============================================
    // State
    // ============================================
    let debounceTimer = null;
    let currentResults = {
        converted: [],      // Array of {url, siteCollection, isError, errorMessage}
        unique: [],         // Array of {siteCollection, count} objects (sorted)
        uniqueInputUrls: 0  // Count of unique input URLs
    };

    // ============================================
    // URL Parsing Logic
    // ============================================

    /**
     * Validates if a URL is a proper SharePoint URL
     * @param {string} url - The URL to validate
     * @returns {{isValid: boolean, errorMessage: string|null}}
     */
    function validateSharePointUrl(url) {
        // Check for basic URL structure
        if (!url.startsWith('http://') && !url.startsWith('https://')) {
            return { isValid: false, errorMessage: 'Missing protocol (http/https)' };
        }

        // Check for sharepoint.com domain (including -my for OneDrive)
        const spDomainRegex = /\.sharepoint\.com($|\/)/i;
        if (!spDomainRegex.test(url)) {
            // Check for common typos
            if (/sharepointcom/i.test(url) || /sharepoint\.co($|\/)/i.test(url)) {
                return { isValid: false, errorMessage: 'Typo in domain (missing dot or incomplete)' };
            }
            return { isValid: false, errorMessage: 'Not a SharePoint URL' };
        }

        // Check for malformed URLs
        try {
            new URL(url);
        } catch {
            return { isValid: false, errorMessage: 'Malformed URL' };
        }

        return { isValid: true, errorMessage: null };
    }

    /**
     * Extracts the site collection URL from a SharePoint URL
     * @param {string} url - The full SharePoint URL
     * @returns {{siteCollection: string, isError: boolean, errorMessage: string|null}}
     */
    function extractSiteCollection(url) {
        // Trim whitespace and trailing slashes
        url = url.trim().replace(/\/+$/, '');

        if (!url) {
            return { siteCollection: null, isError: true, errorMessage: 'Empty URL' };
        }

        // Validate the URL first
        const validation = validateSharePointUrl(url);
        if (!validation.isValid) {
            return { siteCollection: url, isError: true, errorMessage: validation.errorMessage };
        }

        try {
            const urlObj = new URL(url);
            const pathname = urlObj.pathname;
            const origin = urlObj.origin;

            // Pattern: /sites/sitename/...
            const sitesMatch = pathname.match(/^\/sites\/([^\/]+)/i);
            if (sitesMatch) {
                return {
                    siteCollection: `${origin}/sites/${sitesMatch[1]}`,
                    isError: false,
                    errorMessage: null
                };
            }

            // Pattern: /teams/teamname/...
            const teamsMatch = pathname.match(/^\/teams\/([^\/]+)/i);
            if (teamsMatch) {
                return {
                    siteCollection: `${origin}/teams/${teamsMatch[1]}`,
                    isError: false,
                    errorMessage: null
                };
            }

            // Pattern: /personal/username/... (OneDrive)
            const personalMatch = pathname.match(/^\/personal\/([^\/]+)/i);
            if (personalMatch) {
                return {
                    siteCollection: `${origin}/personal/${personalMatch[1]}`,
                    isError: false,
                    errorMessage: null
                };
            }

            // Root site collection (no /sites/, /teams/, or /personal/)
            // This handles URLs like https://tenant.sharepoint.com or https://tenant.sharepoint.com/SitePages/Home.aspx
            return {
                siteCollection: origin,
                isError: false,
                errorMessage: null
            };

        } catch (e) {
            return { siteCollection: url, isError: true, errorMessage: 'Failed to parse URL' };
        }
    }

    // ============================================
    // Processing Logic
    // ============================================

    /**
     * Processes URLs in batches for performance
     * @param {string[]} urls - Array of URLs to process
     * @param {function} onProgress - Progress callback (processed, total)
     * @returns {Promise<{converted: Array, unique: Array, uniqueInputUrls: number}>}
     */
    async function processUrlsInBatches(urls, onProgress) {
        const converted = [];
        const siteCollectionCounts = new Map(); // Track count per site collection
        const uniqueInputSet = new Set();       // Track unique input URLs
        const total = urls.length;

        for (let i = 0; i < total; i += CONFIG.BATCH_SIZE) {
            const batch = urls.slice(i, i + CONFIG.BATCH_SIZE);

            // Process batch
            for (const url of batch) {
                // Track unique input URLs
                uniqueInputSet.add(url.toLowerCase());

                const result = extractSiteCollection(url);
                converted.push({
                    original: url,
                    siteCollection: result.siteCollection,
                    isError: result.isError,
                    errorMessage: result.errorMessage
                });

                if (!result.isError && result.siteCollection) {
                    // Increment count for this site collection
                    const currentCount = siteCollectionCounts.get(result.siteCollection) || 0;
                    siteCollectionCounts.set(result.siteCollection, currentCount + 1);
                }
            }

            // Report progress
            const processed = Math.min(i + CONFIG.BATCH_SIZE, total);
            onProgress(processed, total);

            // Yield to allow UI updates for large datasets
            if (i + CONFIG.BATCH_SIZE < total) {
                await new Promise(resolve => {
                    if (typeof requestIdleCallback !== 'undefined') {
                        requestIdleCallback(resolve, { timeout: 50 });
                    } else {
                        setTimeout(resolve, 0);
                    }
                });
            }
        }

        // Convert map to array and sort alphabetically (case-insensitive)
        const unique = Array.from(siteCollectionCounts.entries())
            .map(([siteCollection, count]) => ({ siteCollection, count }))
            .sort((a, b) => a.siteCollection.toLowerCase().localeCompare(b.siteCollection.toLowerCase()));

        return { converted, unique, uniqueInputUrls: uniqueInputSet.size };
    }

    // ============================================
    // UI Rendering
    // ============================================

    /**
     * Creates an empty state element
     * @param {string} message - Message to display
     * @returns {HTMLElement}
     */
    function createEmptyState(message) {
        const div = document.createElement('div');
        div.className = 'empty-state';
        div.innerHTML = `
            <svg width="48" height="48" viewBox="0 0 48 48" fill="currentColor" opacity="0.3">
                <path d="M24 4C12.95 4 4 12.95 4 24s8.95 20 20 20 20-8.95 20-20S35.05 4 24 4zm0 36c-8.82 0-16-7.18-16-16S15.18 8 24 8s16 7.18 16 16-7.18 16-16 16zm-2-22h4v12h-4V18zm0 16h4v4h-4v-4z"/>
            </svg>
            <p>${message}</p>
        `;
        return div;
    }

    /**
     * Renders the converted URLs to the textarea
     * @param {Array} converted - Array of converted URL objects
     */
    function renderConvertedList(converted) {
        if (converted.length === 0) {
            elements.convertedOutput.value = '';
            return;
        }

        // Build the output string - prefix broken URLs with [INVALID]
        const lines = converted.map(item => {
            if (item.isError) {
                return `[INVALID] ${item.siteCollection || item.original}`;
            }
            return item.siteCollection;
        });

        elements.convertedOutput.value = lines.join('\n');
    }

    /**
     * Renders the unique site collections list
     * @param {Array<{siteCollection: string, count: number}>} unique - Array of unique site collection objects
     */
    function renderUniqueList(unique) {
        const container = elements.uniqueList;
        container.innerHTML = '';

        if (unique.length === 0) {
            container.appendChild(createEmptyState('Unique site collections will appear here, sorted alphabetically'));
            return;
        }

        // Use DocumentFragment for better performance
        const fragment = document.createDocumentFragment();

        for (const item of unique) {
            const div = document.createElement('div');
            div.className = 'url-item url-item-with-count';

            const urlSpan = document.createElement('span');
            urlSpan.className = 'url-text';
            urlSpan.textContent = item.siteCollection;

            const countSpan = document.createElement('span');
            countSpan.className = 'url-row-count';
            countSpan.textContent = item.count.toLocaleString();
            countSpan.title = `${item.count.toLocaleString()} URL${item.count !== 1 ? 's' : ''} map to this site collection`;

            div.appendChild(urlSpan);
            div.appendChild(countSpan);
            fragment.appendChild(div);
        }

        container.appendChild(fragment);
    }

    /**
     * Updates the count badges with animation
     * @param {HTMLElement} badge - The badge element
     * @param {number} count - The new count
     */
    function updateCountBadge(badge, count) {
        badge.textContent = count.toLocaleString();
        badge.classList.remove('updated');
        // Trigger reflow to restart animation
        void badge.offsetWidth;
        badge.classList.add('updated');
    }

    /**
     * Shows/hides progress bar
     * @param {boolean} show - Whether to show the progress bar
     * @param {number} progress - Progress percentage (0-100)
     * @param {string} text - Progress text
     */
    function showProgress(show, progress = 0, text = '') {
        if (show) {
            elements.progressBar.classList.add('active');
            elements.progressFill.style.width = `${progress}%`;
            elements.progressText.textContent = text;
        } else {
            elements.progressBar.classList.remove('active');
        }
    }

    /**
     * Shows a toast notification
     * @param {string} message - Message to display
     */
    function showToast(message) {
        elements.toastMessage.textContent = message;
        elements.toast.classList.add('show');

        setTimeout(() => {
            elements.toast.classList.remove('show');
        }, CONFIG.TOAST_DURATION);
    }

    // ============================================
    // Event Handlers
    // ============================================

    /**
     * Handles input changes with debouncing
     */
    function handleInputChange() {
        // Clear existing timer
        if (debounceTimer) {
            clearTimeout(debounceTimer);
        }

        // Update input count immediately
        const lines = elements.urlInput.value.split('\n').filter(line => line.trim());
        elements.inputCount.textContent = `${lines.length.toLocaleString()} URLs`;

        // Debounce the processing
        debounceTimer = setTimeout(() => {
            processInput();
        }, CONFIG.DEBOUNCE_DELAY);
    }

    /**
     * Processes the input URLs
     */
    async function processInput() {
        const input = elements.urlInput.value;
        const urls = input.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0);

        if (urls.length === 0) {
            currentResults = { converted: [], unique: [], uniqueInputUrls: 0 };
            renderConvertedList([]);
            renderUniqueList([]);
            updateCountBadge(elements.convertedCount, 0);
            updateCountBadge(elements.uniqueCount, 0);
            updateUniqueInputCount(0);
            showProgress(false);
            return;
        }

        // Show progress for large datasets
        const showProgressBar = urls.length > CONFIG.PROGRESS_THRESHOLD;

        if (showProgressBar) {
            showProgress(true, 0, `Processing 0 of ${urls.length.toLocaleString()} URLs...`);
        }

        // Process URLs
        const results = await processUrlsInBatches(urls, (processed, total) => {
            if (showProgressBar) {
                const percent = Math.round((processed / total) * 100);
                showProgress(true, percent, `Processing ${processed.toLocaleString()} of ${total.toLocaleString()} URLs...`);
            }
        });

        // Hide progress
        showProgress(false);

        // Store results
        currentResults = results;

        // Render results
        renderConvertedList(results.converted);
        renderUniqueList(results.unique);
        updateCountBadge(elements.convertedCount, results.converted.length);
        updateCountBadge(elements.uniqueCount, results.unique.length);
        updateUniqueInputCount(results.uniqueInputUrls);
    }

    /**
     * Updates the unique input URLs count display
     * @param {number} count - Number of unique input URLs
     */
    function updateUniqueInputCount(count) {
        const countEl = document.getElementById('uniqueInputCount');
        if (countEl) {
            countEl.textContent = count.toLocaleString();
            countEl.classList.remove('updated');
            void countEl.offsetWidth;
            countEl.classList.add('updated');
        }
    }

    /**
     * Handles clear button click
     */
    function handleClear() {
        elements.urlInput.value = '';
        elements.convertedOutput.value = '';
        currentResults = { converted: [], unique: [], uniqueInputUrls: 0 };
        renderUniqueList([]);
        updateCountBadge(elements.convertedCount, 0);
        updateCountBadge(elements.uniqueCount, 0);
        updateUniqueInputCount(0);
        elements.inputCount.textContent = '0 URLs';
        elements.urlInput.focus();
    }

    /**
     * Copies text to clipboard
     * @param {string} text - Text to copy
     * @param {HTMLElement} button - The button element (for animation)
     */
    async function copyToClipboard(text, button) {
        try {
            await navigator.clipboard.writeText(text);
            button.classList.add('copied');
            showToast('Copied to clipboard!');

            setTimeout(() => {
                button.classList.remove('copied');
            }, 1000);
        } catch (err) {
            // Fallback for older browsers
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();

            try {
                document.execCommand('copy');
                button.classList.add('copied');
                showToast('Copied to clipboard!');

                setTimeout(() => {
                    button.classList.remove('copied');
                }, 1000);
            } catch {
                showToast('Failed to copy');
            }

            document.body.removeChild(textarea);
        }
    }

    /**
     * Handles copy converted URLs
     */
    function handleCopyConverted() {
        const text = elements.convertedOutput.value.trim();
        if (!text) {
            showToast('No URLs to copy');
            return;
        }

        copyToClipboard(text, elements.copyConverted);
    }

    /**
     * Handles copy unique site collections
     */
    function handleCopyUnique() {
        if (currentResults.unique.length === 0) {
            showToast('No URLs to copy');
            return;
        }

        const text = currentResults.unique.map(item => item.siteCollection).join('\n');
        copyToClipboard(text, elements.copyUnique);
    }

    // ============================================
    // Initialization
    // ============================================

    function init() {
        // Attach event listeners
        elements.urlInput.addEventListener('input', handleInputChange);
        elements.clearBtn.addEventListener('click', handleClear);
        elements.copyConverted.addEventListener('click', handleCopyConverted);
        elements.copyUnique.addEventListener('click', handleCopyUnique);

        // Handle paste event for immediate feedback
        elements.urlInput.addEventListener('paste', () => {
            // Small delay to ensure pasted content is in the textarea
            setTimeout(handleInputChange, 10);
        });

        // Focus input on load
        elements.urlInput.focus();
    }

    // Start the app
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
