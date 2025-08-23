/**
 * XDR Story Parser - Main JavaScript Module
 * Handles file upload, parsing, and process tree visualization
 */

class XDRTreeVisualizer {
    constructor() {
        this.data = null;
        this.originalData = null;
        this.isAnonymized = false;
        this.isZoomedMode = false;
        this.zoomedNodeId = null;
        this.stats = {
            total: 0,
            processes: 0,
            files: 0,
            accounts: 0,
            networks: 0,
            registry: 0,
            others: 0
        };
        this.expandedNodes = new Set();
        this.anonymizationInfo = {
            usernames: new Set(),
            domains: new Set(),
            deviceIds: new Set(),
            deviceNames: new Set(),
            sids: new Set()
        };
        this.initializeEventListeners();
        this.initializeTheme();
    }

    /**
     * Initialize all event listeners for the application
     */
    initializeEventListeners() {
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');

        if (!uploadArea || !fileInput) {
            console.error('Required DOM elements not found');
            return;
        }

        // File input change event
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFile(e.target.files[0]);
            }
        });

        // Drag and drop events
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();

            if (uploadArea.classList.contains('minimized')) {
                this.restoreUploadSection();
            }

            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });

        // Click to upload
        uploadArea.addEventListener('click', (e) => {
            if (e.target.tagName !== 'BUTTON') {
                if (uploadArea.classList.contains('minimized')) {
                    this.restoreUploadSection();
                    return;
                }
                fileInput.click();
            }
        });

        // Keyboard accessibility
        uploadArea.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                fileInput.click();
            }
        });

        // Anonymization toggle
        const anonymizeCheckbox = document.getElementById('anonymize-checkbox');
        if (anonymizeCheckbox) {
            anonymizeCheckbox.addEventListener('change', (e) => {
                this.toggleAnonymization(e.target.checked);
            });
        }
    }

    /**
     * Handle file selection and processing
     * @param {File} file - The selected file
     */
    async handleFile(file) {
        // Validate file type
        if (!this.isValidFileType(file.name)) {
            this.showError('Please select a valid JSON or JSONC file.');
            return;
        }

        // Validate file size (max 50MB)
        if (file.size > 50 * 1024 * 1024) {
            this.showError('File size too large. Please select a file smaller than 50MB.');
            return;
        }

        this.showLoading();

        try {
            const text = await this.readFile(file);

            // Clean JSONC content first
            let jsonText = this.fixForwardSlashes(text);

            // Try parsing with basic cleaning first
            try {
                this.originalData = JSON.parse(jsonText);
                this.data = JSON.parse(jsonText);
            } catch (parseError) {
                console.warn('Basic parsing failed, applying forward slash fix:', parseError.message);

                // Simple fix: escape forward slashes in JSON string values
                jsonText = this.fixForwardSlashes(jsonText);

                try {
                    this.originalData = JSON.parse(jsonText);
                    this.data = JSON.parse(jsonText);
                } catch (secondParseError) {
                    console.warn('Forward slash fix failed, attempting deep cleaning:', secondParseError.message);

                    // Fallback to existing deep cleaning
                    jsonText = this.deepCleanJson(jsonText);

                    try {
                        this.originalData = JSON.parse(jsonText);
                        this.data = JSON.parse(jsonText);
                    } catch (thirdParseError) {
                        console.warn('Deep cleaning failed, attempting JSON repair:', thirdParseError.message);

                        // Last resort: try to repair the JSON structure
                        jsonText = this.repairJson(jsonText);
                        this.originalData = JSON.parse(jsonText);
                        this.data = JSON.parse(jsonText);
                    }
                }
            }

            this.validateDataStructure();
            this.extractAnonymizationInfo();
            this.processData();
            this.renderTree();

        } catch (error) {
            console.error('Error processing file:', error);
            this.showError(`Error parsing file: ${error.message}`);
        }
    }

    /**
     * Fix forward slashes by escaping them throughout the JSON
     * @param {string} jsonText - JSON text that may contain unescaped forward slashes
     * @returns {string} - JSON text with all forward slashes escaped
     */
    fixForwardSlashes(jsonText) {
        console.log('Fixing forward slashes by escaping all / to \/...');

        // Global replacement: / -> \/
        let fixed = jsonText.replace(/\//g, '\\/');

        console.log('Forward slash fix complete');
        return fixed;
    }

    /**
     * Check if file type is valid
     * @param {string} filename - The filename to check
     * @returns {boolean} - Whether the file type is valid
     */
    isValidFileType(filename) {
        const validExtensions = ['.json', '.jsonc'];
        return validExtensions.some(ext => filename.toLowerCase().endsWith(ext));
    }

    /**
     * Read file contents
     * @param {File} file - The file to read
     * @returns {Promise<string>} - The file contents
     */
    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsText(file);
        });
    }

    /**
     * Clean JSONC content by removing comments, trailing commas, and handling control characters
     * @param {string} text - The JSONC text to clean
     * @returns {string} - Clean JSON text
     */
    cleanJsonC(text) {
        console.log('Starting conservative JSON cleaning...');

        // Basic JSONC cleaning
        let cleaned = text
            .replace(/\/\*[\s\S]*?\*\//g, '') // Remove /* */ comments
            .replace(/\/\/.*$/gm, '') // Remove // comments
            .replace(/,(\s*[}\]])/g, '$1'); // Remove trailing commas

        console.log('JSON cleaning completed');
        return cleaned;
    }

    /**
     * Deep clean JSON with more aggressive fixes for problematic content
     * @param {string} text - The JSON text to deep clean
     * @returns {string} - Deeply cleaned JSON text
     */
    deepCleanJson(text) {
        let cleaned = text;

        try {
            console.log('Starting conservative deep JSON cleaning...');

            // Handle structural issues
            cleaned = cleaned
                .replace(/,(\s*[}\]])/g, '$1') // Remove trailing commas
                .replace(/([{\[])\s*,/g, '$1') // Remove commas right after opening brackets
                .trim();

        } catch (cleaningError) {
            console.warn('Deep cleaning encountered issues:', cleaningError);
        }

        return cleaned;
    }

    /**
     * Last resort JSON repair for severely malformed JSON
     * @param {string} text - The malformed JSON text
     * @returns {string} - Repaired JSON text
     */
    repairJson(text) {
        console.log('Attempting JSON repair...');

        try {
            // Try to extract and repair the main structure
            let repaired = text.trim();

            // Find the main JSON object boundaries more carefully
            let depth = 0;
            let start = -1;
            let end = -1;
            let inString = false;
            let escape = false;

            for (let i = 0; i < repaired.length; i++) {
                const char = repaired[i];

                if (escape) {
                    escape = false;
                    continue;
                }

                if (char === '\\') {
                    escape = true;
                    continue;
                }

                if (char === '"' && !escape) {
                    inString = !inString;
                    continue;
                }

                if (!inString) {
                    if (char === '{') {
                        if (start === -1) start = i;
                        depth++;
                    } else if (char === '}') {
                        depth--;
                        if (depth === 0 && start !== -1) {
                            end = i;
                            break;
                        }
                    }
                }
            }

            if (start !== -1 && end !== -1) {
                repaired = repaired.substring(start, end + 1);
                console.log('Extracted main JSON object');
            }

            repaired = repaired
                // Fix incomplete key-value pairs
                .replace(/:\s*$/gm, ': ""')
                .replace(/:\s*,/g, ': "",')
                .replace(/:\s*}/g, ': ""}')
                .replace(/:\s*]/g, ': ""]')
                // Fix incomplete arrays and objects
                .replace(/,\s*}/g, '}')
                .replace(/,\s*]/g, ']')
                // Fix missing quotes on keys
                .replace(/([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:/g, '$1"$2":')
                // Remove trailing commas
                .replace(/,(\s*[}\]])/g, '$1');

            console.log('JSON repair completed');
            return repaired;

        } catch (repairError) {
            console.error('JSON repair failed:', repairError);
            // Ultimate fallback: return a minimal valid JSON
            return '{"error": "Failed to parse malformed JSON", "items": []}';
        }
    }

    /**
     * Extract information that should be anonymized
     */
    extractAnonymizationInfo() {
        this.anonymizationInfo = {
            usernames: new Set(),
            domains: new Set(),
            deviceIds: new Set(),
            deviceNames: new Set(),
            sids: new Set()
        };

        // Extract from main user and device info
        if (this.data.mainUser) {
            if (this.data.mainUser.name && !this.isSystemAccount(this.data.mainUser.name)) {
                this.anonymizationInfo.usernames.add(this.data.mainUser.name);
            }
            if (this.data.mainUser.domainName && !this.isSystemDomain(this.data.mainUser.domainName)) {
                this.anonymizationInfo.domains.add(this.data.mainUser.domainName);
            }
            if (this.data.mainUser.sid) this.anonymizationInfo.sids.add(this.data.mainUser.sid);
        }

        if (this.data.deviceId) this.anonymizationInfo.deviceIds.add(this.data.deviceId);
        if (this.data.deviceName) {
            // Add the full device name for redaction
            this.anonymizationInfo.deviceNames.add(this.data.deviceName);

            // For FQDN device names, also add just the hostname part
            const deviceParts = this.data.deviceName.split('.');
            if (deviceParts.length > 1) {
                // Add just the hostname (first part) for separate redaction
                const hostname = deviceParts[0];
                this.anonymizationInfo.deviceNames.add(hostname);

                // Domain extraction - only add proper domains, not infrastructure components
                if (deviceParts.length >= 3) {
                    const potentialDomain = deviceParts.slice(-2).join('.');
                    // Only add if it looks like a proper domain and isn't a system domain
                    if (potentialDomain.match(/^[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$/) &&
                        !this.isSystemDomain(potentialDomain)) {
                        this.anonymizationInfo.domains.add(potentialDomain);
                    }
                }
            }
        }

        // Extract from all items recursively
        if (this.data.items) {
            this.data.items.forEach(item => this.extractItemAnonymizationInfo(item));
        }
    }

    /**
     * Extract anonymization info from individual items
     */
    extractItemAnonymizationInfo(item) {
        if (!item) return;

        // Extract from entity
        if (item.entity) {
            if (item.entity.User) {
                // Skip system usernames
                if (item.entity.User.UserName && !this.isSystemAccount(item.entity.User.UserName)) {
                    this.anonymizationInfo.usernames.add(item.entity.User.UserName);
                }
                // Skip system domains
                if (item.entity.User.DomainName && !this.isSystemDomain(item.entity.User.DomainName)) {
                    this.anonymizationInfo.domains.add(item.entity.User.DomainName);
                }
                if (item.entity.User.Sid) this.anonymizationInfo.sids.add(item.entity.User.Sid);
            }
        }

        // Process children and nested items
        if (item.children) {
            item.children.forEach(child => this.extractItemAnonymizationInfo(child));
        }
        if (item.nestedItems) {
            item.nestedItems.forEach(nested => this.extractItemAnonymizationInfo(nested));
        }
    }

    /**
     * Check if a username is a system account that shouldn't be redacted
     */
    isSystemAccount(username) {
        const systemAccounts = [
            'SYSTEM',
            'LOCAL SERVICE',
            'NETWORK SERVICE',
            'ANONYMOUS LOGON',
            'SERVICE',
            'BATCH',
            'DIALUP',
            'EVERYONE',
            'AUTHENTICATED USERS',
            'IUSR',
            'IWAM',
            'ASPNET',
            'KRBTGT',
            'GUEST'
        ];
        return systemAccounts.includes(username.toUpperCase());
    }

    /**
     * Check if a domain name is a system domain that shouldn't be redacted
     */
    isSystemDomain(domainName) {
        const systemDomains = [
            'NT AUTHORITY',
            'NT SERVICE',
            'BUILTIN'
        ];
        return systemDomains.includes(domainName.toUpperCase());
    }

    /**
     * Toggle anonymization on/off
     */
    toggleAnonymization(enable) {
        this.isAnonymized = enable;

        if (enable) {
            this.data = this.createAnonymizedData(JSON.parse(JSON.stringify(this.originalData)));
        } else {
            this.data = JSON.parse(JSON.stringify(this.originalData));
        }

        this.updateInvestigationInfo();
        this.renderTree();
    }

    /**
     * Create anonymized version of the data
     */
    createAnonymizedData(data) {
        const anonymized = JSON.parse(JSON.stringify(data));

        // Recursively anonymize all string values in the entire JSON structure
        this.deepAnonymizeObject(anonymized);

        return anonymized;
    }

    /**
     * Recursively anonymize all string values in an object/array
     */
    deepAnonymizeObject(obj) {
        if (obj === null || obj === undefined) return;

        if (typeof obj === 'string') {
            return this.anonymizeString(obj);
        }

        if (Array.isArray(obj)) {
            for (let i = 0; i < obj.length; i++) {
                if (typeof obj[i] === 'string') {
                    obj[i] = this.anonymizeString(obj[i]);
                } else if (typeof obj[i] === 'object') {
                    this.deepAnonymizeObject(obj[i]);
                }
            }
        } else if (typeof obj === 'object') {
            for (const key in obj) {
                if (obj.hasOwnProperty(key)) {
                    if (typeof obj[key] === 'string') {
                        obj[key] = this.anonymizeString(obj[key]);
                    } else if (typeof obj[key] === 'object') {
                        this.deepAnonymizeObject(obj[key]);
                    }
                }
            }
        }
    }

    /**
     * Anonymize individual item (legacy method - now uses deepAnonymizeObject)
     */
    anonymizeItem(item) {
        if (!item) return;
        this.deepAnonymizeObject(item);
    }

    /**
     * Anonymize a string by replacing sensitive information
     */
    anonymizeString(str) {
        let result = str;

        // Replace device names first (before domains) to handle FQDNs properly
        this.anonymizationInfo.deviceNames.forEach(deviceName => {
            const regex = new RegExp(this.escapeRegExp(deviceName), 'gi');
            result = result.replace(regex, 'REDACTED');
        });

        // Replace usernames
        this.anonymizationInfo.usernames.forEach(username => {
            const regex = new RegExp(this.escapeRegExp(username), 'gi');
            result = result.replace(regex, 'REDACTED');
        });

        // Replace domains - use word boundaries for short domains to prevent partial matches
        this.anonymizationInfo.domains.forEach(domain => {
            // For very short domain components (3 chars or less), use word boundaries
            if (domain.length <= 3) {
                const regex = new RegExp('\\b' + this.escapeRegExp(domain) + '\\b', 'gi');
                result = result.replace(regex, 'REDACTED');
            } else {
                const regex = new RegExp(this.escapeRegExp(domain), 'gi');
                result = result.replace(regex, 'REDACTED');
            }
        });

        // Replace device IDs
        this.anonymizationInfo.deviceIds.forEach(deviceId => {
            const regex = new RegExp(this.escapeRegExp(deviceId), 'gi');
            result = result.replace(regex, 'REDACTED');
        });

        // Replace SIDs
        this.anonymizationInfo.sids.forEach(sid => {
            const regex = new RegExp(this.escapeRegExp(sid), 'gi');
            result = result.replace(regex, 'REDACTED');
        });

        return result;
    }

    /**
     * Escape special regex characters
     */
    escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Update the investigation info display
     */
    updateInvestigationInfo() {
        const investigationInfo = document.getElementById('investigation-info');

        if (this.data && (this.data.mainUser || this.data.deviceId || this.data.deviceName)) {
            investigationInfo.style.display = 'block';

            // Main user info
            const mainUser = this.data.mainUser;
            const userName = document.getElementById('main-user-name');
            const userDomain = document.getElementById('main-user-domain');
            const userSid = document.getElementById('main-user-sid');

            if (userName) userName.textContent = mainUser?.name || '-';
            if (userDomain) userDomain.textContent = mainUser?.domainName || '-';
            if (userSid) userSid.textContent = mainUser?.sid || '-';

            // Device info
            const deviceName = document.getElementById('device-name');
            const deviceId = document.getElementById('device-id');

            if (deviceName) deviceName.textContent = this.data.deviceName || '-';
            if (deviceId) deviceId.textContent = this.data.deviceId || '-';

            // Add redacted styling if anonymized
            const valueElements = investigationInfo.querySelectorAll('.info-value');
            valueElements.forEach(el => {
                if (this.isAnonymized && el.textContent === 'REDACTED') {
                    el.classList.add('redacted');
                } else {
                    el.classList.remove('redacted');
                }
            });
        } else {
            investigationInfo.style.display = 'none';
        }
    }

    /**
     * Validate that the data structure matches expected XDR story format
     */
    validateDataStructure() {
        if (!this.data) {
            throw new Error('No data found in file');
        }

        if (!this.data.items || !Array.isArray(this.data.items)) {
            throw new Error('Invalid data structure: missing or invalid "items" array');
        }

        if (this.data.items.length === 0) {
            throw new Error('No items found in the data');
        }
    }

    /**
     * Process the data and calculate statistics
     */
    processData() {
        // Reset statistics
        this.stats = {
            total: 0,
            processes: 0,
            files: 0,
            accounts: 0,
            networks: 0,
            registry: 0,
            others: 0
        };

        // Process all items recursively
        if (this.data && this.data.items) {
            this.data.items.forEach(item => this.countItems(item));
        }

        console.log('Data processing complete:', this.stats);
    }

    /**
     * Recursively count items by type
     * @param {Object} item - The item to count
     */
    countItems(item) {
        if (!item) return;

        this.stats.total++;

        // Determine item type
        const type = this.getItemType(item);
        switch (type) {
            case 'process':
                this.stats.processes++;
                break;
            case 'file':
                this.stats.files++;
                break;
            case 'account':
                this.stats.accounts++;
                break;
            case 'network':
                this.stats.networks++;
                break;
            case 'registry':
                this.stats.registry++;
                break;
            default:
                this.stats.others++;
        }

        // Process children recursively
        if (item.children && Array.isArray(item.children)) {
            item.children.forEach(child => this.countItems(child));
        }

        // Process nested items
        if (item.nestedItems && Array.isArray(item.nestedItems)) {
            item.nestedItems.forEach(nested => this.countItems(nested));
        }
    }

    /**
     * Determine the type of an item
     * @param {Object} item - The item to analyze
     * @returns {string} - The item type
     */
    getItemType(item) {
        return item.type || item.actionType || 'other';
    }

    /**
     * Render the complete tree visualization
     */
    renderTree() {
        const treeContainer = document.getElementById('tree-content');
        const processTree = document.getElementById('process-tree');

        if (!treeContainer || !processTree) {
            console.error('Required DOM elements not found');
            return;
        }

        // Update statistics display
        this.updateStatsDisplay();

        // Update investigation info
        this.updateInvestigationInfo();

        // Clear previous content
        treeContainer.innerHTML = '';

        // Render tree nodes - all children will be shown by default
        if (this.data && this.data.items) {
            const fragment = document.createDocumentFragment();
            this.data.items.forEach(item => {
                this.renderNode(item, fragment, 0);
            });
            treeContainer.appendChild(fragment);
        }

        // Show the tree container
        processTree.style.display = 'block';

        // Show the type legend
        const typeLegend = document.querySelector('.type-legend');
        if (typeLegend) {
            typeLegend.style.display = 'block';
        }

        // Show the tree visualization
        const treeVisualization = document.querySelector('.tree-visualization');
        if (treeVisualization) {
            treeVisualization.style.display = 'block';
        }

        // Show the download button
        const downloadBtn = document.getElementById('download-json-btn');
        if (downloadBtn) {
            downloadBtn.style.display = 'inline-block';
        }

        // Show the analysis tools section
        const analysisSection = document.getElementById('analysis-tools');
        if (analysisSection) {
            analysisSection.style.display = 'block';
        }

        // Minimize the upload section
        this.minimizeUploadSection();

        // Scroll to tree with smooth animation
        setTimeout(() => {
            processTree.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
    }

    /**
     * Update the statistics display
     */
    updateStatsDisplay() {
        const elements = {
            'total-items': this.stats.total,
            'process-count': this.stats.processes,
            'file-count': this.stats.files
        };

        Object.entries(elements).forEach(([id, value]) => {
            const element = document.getElementById(id);
            if (element) {
                element.textContent = value.toLocaleString();
            }
        });
    }

    /**
     * Get a consistent node ID
     * @param {Object} node - The node
     * @returns {string} - Consistent node ID
     */
    getNodeId(node) {
        return node.id || `node-${node.title?.main || 'unknown'}-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;
    }

    /**
     * Render a single node and its children
     * @param {Object} node - The node to render
     * @param {DocumentFragment|HTMLElement} container - The container to append to
     * @param {number} level - The nesting level
     */
    renderNode(node, container, level) {
        if (!node) return;

        // Filter out specific node types we don't want to display
        const nodeTitle = this.getNodeTitle(node);
        const nodeSubtitle = this.getNodeSubtitle(node);

        // Skip nodes with "PE metadata" or "User" in their intro/subtitle
        if (nodeSubtitle && (nodeSubtitle.includes('PE metadata') || nodeSubtitle.includes('User') || nodeSubtitle.includes('Web data file'))) {
            // Still process children if any
            if (node.children && Array.isArray(node.children)) {
                node.children.forEach(child => {
                    this.renderNode(child, container, level);
                });
            }
            if (node.nestedItems && Array.isArray(node.nestedItems)) {
                node.nestedItems.forEach(nested => {
                    this.renderNode(nested, container, level);
                });
            }
            return;
        }

        const nodeDiv = document.createElement('div');
        nodeDiv.className = 'tree-node';
        const nodeId = this.getNodeId(node);
        nodeDiv.dataset.nodeId = nodeId;

        // Create tree structure visualization
        const indent = this.createIndentation(level);

        // Get node information
        const type = this.getItemType(node);
        const icon = this.getNodeIcon(type, node);
        const title = this.getNodeTitle(node);
        const subtitle = this.getNodeSubtitle(node);
        const commandLine = this.getNodeCommandLine(node);
        const time = this.formatTime(node.time);

        // Check if this node or any descendants have alerts
        const hasAlertsInTree = this.nodeHasAlertsInTree(node);

        // Check if node has details to show
        const hasDetails = this.nodeHasDetails(node);

        // Check if node has any children (both direct children and nested items)
        const hasChildren = (node.children && Array.isArray(node.children) && node.children.length > 0);
        const hasNestedItems = (node.nestedItems && Array.isArray(node.nestedItems) && node.nestedItems.length > 0);
        const hasAnyChildren = hasChildren || hasNestedItems;

        // Build expand button for any children
        const expandButton = hasAnyChildren ?
            `<span class="expand-button" onclick="xdrVisualizer.toggleNodeChildren('${nodeId}')" title="Click to expand/collapse children">▼</span>` :
            '<span class="expand-placeholder"></span>';

        // Build node HTML
        nodeDiv.innerHTML = `
            <div class="node-buttons">
            <span class="tree-indent">${indent}</span>
            ${expandButton}
            <span class="tree-icon">${icon}</span>
            ${hasAnyChildren ? `<span class="zoom-button" onclick="xdrVisualizer.zoomToNode('${nodeId}')" title="Zoom to this node and its children">🔍</span>` : '<span class="zoom-placeholder"></span>'}
            </div>
            <div class="node-content ${type}" ${hasDetails ? `onclick="xdrVisualizer.toggleNodeDetails('${nodeId}')"` : ''} title="${hasDetails ? 'Click for details' : ''}">
                <div class="node-title-row">
                    <div class="node-title">
                        ${hasAlertsInTree ? '<span class="alert-indicator">🚨</span>' : ''}
                        ${this.escapeHtml(title)}
                    </div>
                    ${time ? `<div class="node-time">${time}</div>` : ''}
                </div>
                ${subtitle ? `<div class="node-subtitle">${this.escapeHtml(subtitle)}</div>` : ''}
                ${commandLine ? `<div class="node-commandline">${this.escapeHtml(this.unescapeForwardSlashes(commandLine))}</div>` : ''}
            </div>
        `;

        // Add details panel if there are details to show
        if (hasDetails) {
            const detailsPanel = document.createElement('div');
            detailsPanel.className = 'details-panel';
            detailsPanel.style.display = 'none'; // Start collapsed
            detailsPanel.innerHTML = this.renderNodeDetails(node);
            nodeDiv.appendChild(detailsPanel);
        }

        container.appendChild(nodeDiv);

        // Add associated alerts as adjacent nodes
        if (node.associatedAlerts && Array.isArray(node.associatedAlerts) && node.associatedAlerts.length > 0) {
            node.associatedAlerts.forEach(alert => {
                if (alert.alertDisplayName) {
                    const alertDiv = document.createElement('div');
                    alertDiv.className = 'tree-node alert-node';

                    const alertNodeId = `alert-${nodeId}-${Math.random().toString(36).substr(2, 5)}`;
                    alertDiv.dataset.nodeId = alertNodeId;

                    alertDiv.innerHTML = `
                        <span class="tree-indent">${this.createIndentation(level)}</span>
                        <span class="expand-placeholder"></span>
                        <span class="tree-icon">🚨</span>
                        <div class="node-content alert">
                            <div class="node-title">${this.escapeHtml(alert.alertDisplayName)}</div>
                        </div>
                    `;

                    container.appendChild(alertDiv);
                }
            });
        }

        // Create children container for all types of children
        if (hasAnyChildren) {
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'children-container';
            childrenContainer.dataset.nodeId = nodeId;
            // By default, nothing is collapsed (as requested)
            childrenContainer.style.display = 'block';

            // Add direct children first
            if (hasChildren) {
                node.children.forEach(child => {
                    this.renderNode(child, childrenContainer, level + 1);
                });
            }

            // Add nested items
            if (hasNestedItems) {
                node.nestedItems.forEach(nested => {
                    this.renderNode(nested, childrenContainer, level + 1);
                });
            }

            container.appendChild(childrenContainer);
        }
    }

    /**
     * Create indentation string for tree structure
     * @param {number} level - The nesting level
     * @returns {string} - The indentation string
     */
    createIndentation(level) {
        if (level === 0) return '';

        let indent = '';
        for (let i = 0; i < level - 1; i++) {
            indent += '    '; // 4 spaces for each level
        }
        indent += '└── '; // Clean L-shaped connector

        return indent;
    }

    /**
     * Check if node has children or nested items
     * @param {Object} node - The node to check
     * @returns {boolean} - Whether the node has children
     */
    nodeHasChildren(node) {
        return (node.children && node.children.length > 0) ||
            (node.nestedItems && node.nestedItems.length > 0);
    }

    /**
     * Check if node has details to display
     * @param {Object} node - The node to check
     * @returns {boolean} - Whether the node has details
     */
    nodeHasDetails(node) {
        const hasNodeDetails = node.details && Array.isArray(node.details) && node.details.length > 0;
        const hasAdditionalDetails = node.additionalDetails && Array.isArray(node.additionalDetails) && node.additionalDetails.length > 0;
        const hasEntityDetails = node.entity && this.entityHasMeaningfulData(node.entity);

        return hasNodeDetails || hasAdditionalDetails || hasEntityDetails;
    }

    /**
     * Check if entity has meaningful data to display
     * @param {Object} entity - The entity to check
     * @returns {boolean} - Whether the entity has meaningful data
     */
    entityHasMeaningfulData(entity) {
        if (!entity) return false;

        // Check for ImageFile information
        if (entity.ImageFile) {
            const img = entity.ImageFile;
            if (img.FullPath || img.Size || img.Sha256 || img.Sha1 || img.Md5 || img.CreationTime) {
                return true;
            }
        }

        // Check for User information
        if (entity.User) {
            const user = entity.User;
            if (user.DomainName || user.UserName || user.Sid) {
                return true;
            }
        }

        // Check for Process information
        if (entity.ProcessId || entity.Commandline || entity.CreatingProcessId ||
            entity.CreatingProcessName || entity.CreationTime || entity.IntegrityLevel ||
            entity.TokenElevation) {
            return true;
        }

        return false;
    }

    /**
     * Check if node or any of its descendants has alerts
     * @param {Object} node - The node to check
     * @returns {boolean} - Whether the node or its descendants have alerts
     */
    nodeHasAlertsInTree(node) {
        // Check if the current node has associated alerts
        if (node.associatedAlerts && node.associatedAlerts.length > 0) {
            return true;
        }

        // Check children recursively
        if (node.children && node.children.length > 0) {
            for (const child of node.children) {
                if (this.nodeHasAlertsInTree(child)) {
                    return true;
                }
            }
        }

        // Check nested items recursively
        if (node.nestedItems && node.nestedItems.length > 0) {
            for (const nested of node.nestedItems) {
                if (this.nodeHasAlertsInTree(nested)) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Toggle children visibility for a node
     * @param {string} nodeId - The node ID
     */
    toggleNodeChildren(nodeId) {
        const childrenContainer = document.querySelector(`.children-container[data-node-id="${nodeId}"]`);
        const expandButton = document.querySelector(`[data-node-id="${nodeId}"] .expand-button`);

        if (!childrenContainer || !expandButton) return;

        if (childrenContainer.style.display === 'none') {
            // Expand
            childrenContainer.style.display = 'block';
            expandButton.textContent = '▼';
            expandButton.title = 'Click to collapse children';
        } else {
            // Collapse
            childrenContainer.style.display = 'none';
            expandButton.textContent = '▶';
            expandButton.title = 'Click to expand children';
        }
    }

    /**
     * Expand all nodes in the tree
     */
    expandAll() {
        console.log('ExpandAll called, isZoomedMode:', this.isZoomedMode, 'zoomedNodeId:', this.zoomedNodeId);

        if (this.isZoomedMode && this.zoomedNodeId) {
            const visibleNodes = document.querySelectorAll('.tree-node');
            const visibleTreeNodes = Array.from(visibleNodes).filter(node => node.offsetParent !== null);

            console.log('Found visible tree nodes:', visibleTreeNodes.length);

            visibleTreeNodes.forEach((node, index) => {
                const nodeId = node.dataset.nodeId;
                const childrenContainer = document.querySelector(`.children-container[data-node-id="${nodeId}"]`);
                const expandButton = node.querySelector('.expand-button');

                console.log(`Node ${index} (${nodeId}):`, {
                    hasContainer: !!childrenContainer,
                    hasButton: !!expandButton,
                    containerVisible: childrenContainer ? childrenContainer.offsetParent !== null : false
                });

                if (childrenContainer) {
                    childrenContainer.style.display = 'block';
                }

                if (expandButton) {
                    expandButton.textContent = '▼';
                    expandButton.title = 'Click to collapse children';
                }
            });
        } else {
            // Normal mode: expand all nodes in the entire tree
            const childrenContainers = document.querySelectorAll('.children-container');
            const expandButtons = document.querySelectorAll('.expand-button');
            console.log('Normal mode - containers:', childrenContainers.length, 'buttons:', expandButtons.length);

            childrenContainers.forEach(container => {
                container.style.display = 'block';
            });

            expandButtons.forEach(button => {
                button.textContent = '▼';
                button.title = 'Click to collapse children';
            });
        }

        console.log('ExpandAll completed');
    }

    /**
     * Collapse all nodes in the tree
     */
    collapseAll() {
        console.log('CollapseAll called, isZoomedMode:', this.isZoomedMode, 'zoomedNodeId:', this.zoomedNodeId);

        if (this.isZoomedMode && this.zoomedNodeId) {
            const visibleNodes = document.querySelectorAll('.tree-node');
            const visibleTreeNodes = Array.from(visibleNodes).filter(node => node.offsetParent !== null);

            console.log('Found visible tree nodes:', visibleTreeNodes.length);

            visibleTreeNodes.forEach((node, index) => {
                const nodeId = node.dataset.nodeId;
                const childrenContainer = document.querySelector(`.children-container[data-node-id="${nodeId}"]`);
                const expandButton = node.querySelector('.expand-button');

                console.log(`Node ${index} (${nodeId}):`, {
                    hasContainer: !!childrenContainer,
                    hasButton: !!expandButton,
                    containerVisible: childrenContainer ? childrenContainer.offsetParent !== null : false
                });

                if (childrenContainer) {
                    childrenContainer.style.display = 'none';
                }

                if (expandButton) {
                    expandButton.textContent = '▶';
                    expandButton.title = 'Click to expand children';
                }
            });
        } else {
            // Normal mode: collapse all nodes in the entire tree
            const childrenContainers = document.querySelectorAll('.children-container');
            const expandButtons = document.querySelectorAll('.expand-button');
            console.log('Normal mode - containers:', childrenContainers.length, 'buttons:', expandButtons.length);

            childrenContainers.forEach(container => {
                container.style.display = 'none';
            });

            expandButtons.forEach(button => {
                button.textContent = '▶';
                button.title = 'Click to expand children';
            });
        }

        console.log('CollapseAll completed');
    }

    /**
     * Zoom to a specific node and hide all others
     * @param {string} nodeId - The node ID to zoom to
     */
    zoomToNode(nodeId) {
        this.isZoomedMode = true;
        this.zoomedNodeId = nodeId;

        // Hide all nodes first
        const allNodes = document.querySelectorAll('.tree-node');
        allNodes.forEach(node => {
            node.style.display = 'none';
        });

        // Find and show the target node and its descendants
        const targetNode = document.querySelector(`[data-node-id="${nodeId}"]`);
        if (targetNode) {
            // Reset the target node as the new root (level 0)
            this.makeNodeRoot(targetNode);
            this.showNodeAndDescendants(targetNode);

            // Expand the target node to show its children
            const childrenContainer = document.querySelector(`.children-container[data-node-id="${nodeId}"]`);
            if (childrenContainer) {
                childrenContainer.style.display = 'block';
                const expandButton = targetNode.querySelector('.expand-button');
                if (expandButton) {
                    expandButton.textContent = '▼';
                    expandButton.title = 'Click to collapse children';
                }

                this.adjustChildrenIndentation(childrenContainer, 1);
            }
        }

        // Add zoom-out button to tree controls
        this.addZoomOutButton();

        // Scroll to the zoomed node
        if (targetNode) {
            setTimeout(() => {
                targetNode.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
        }
    }

    /**
     * Make a node appear as the root node (level 0 indentation)
     * @param {HTMLElement} nodeElement - The node element to make root
     */
    makeNodeRoot(nodeElement) {
        if (!nodeElement) return;

        const indentElement = nodeElement.querySelector('.tree-indent');
        if (indentElement) {
            // Remove all indentation to make it appear as root level
            indentElement.innerHTML = '';
        }

        // Add a visual indicator that this is the zoomed root
        nodeElement.classList.add('zoomed-root');
    }

    /**
     * Recursively adjust indentation of children to start from given level
     * @param {HTMLElement} container - The children container
     * @param {number} startLevel - The starting level for children
     */
    adjustChildrenIndentation(container, startLevel) {
        if (!container) return;

        const childNodes = container.querySelectorAll(':scope > .tree-node');
        childNodes.forEach(childNode => {
            const indentElement = childNode.querySelector('.tree-indent');
            if (indentElement) {
                // Create new indentation for this level
                indentElement.innerHTML = this.createIndentation(startLevel);
            }

            // Process this child's children recursively
            const nodeId = childNode.dataset.nodeId;
            const childContainer = container.querySelector(`.children-container[data-node-id="${nodeId}"]`);
            if (childContainer) {
                this.adjustChildrenIndentation(childContainer, startLevel + 1);
            }
        });
    }

    /**
     * Show a node and all its descendants
     * @param {HTMLElement} nodeElement - The node element to show
     */
    showNodeAndDescendants(nodeElement) {
        if (!nodeElement) return;

        nodeElement.style.display = '';

        // Find the associated children container
        const nodeId = nodeElement.dataset.nodeId;
        const childrenContainer = document.querySelector(`.children-container[data-node-id="${nodeId}"]`);

        if (childrenContainer) {
            childrenContainer.style.display = 'block';

            // Show all child nodes recursively
            const childNodes = childrenContainer.querySelectorAll('.tree-node');
            childNodes.forEach(childNode => {
                this.showNodeAndDescendants(childNode);
            });
        }
    }

    /**
     * Exit zoom mode and show all nodes
     */
    exitZoomMode() {
        this.isZoomedMode = false;
        const previousZoomedNodeId = this.zoomedNodeId;
        this.zoomedNodeId = null;

        // Show all nodes
        const allNodes = document.querySelectorAll('.tree-node');
        allNodes.forEach(node => {
            node.style.display = 'block';
            // Remove zoomed root styling
            node.classList.remove('zoomed-root');
        });

        this.renderTree();

        this.removeZoomOutButton();

        if (previousZoomedNodeId) {
            setTimeout(() => {
                const originalNode = document.querySelector(`[data-node-id="${previousZoomedNodeId}"]`);
                if (originalNode) {
                    originalNode.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            }, 100);
        }
    }

    /**
     * Add zoom-out button to tree controls
     */
    addZoomOutButton() {
        // Remove existing zoom-out button if present
        this.removeZoomOutButton();

        const bottomControls = document.querySelector('.tree-bottom-controls');
        if (bottomControls) {
            const zoomOutBtn = document.createElement('button');
            zoomOutBtn.className = 'btn-small zoom-out-btn';
            zoomOutBtn.id = 'zoom-out-btn';
            zoomOutBtn.innerHTML = '🔍️❌ Exit Zoom';
            zoomOutBtn.onclick = () => this.exitZoomMode();
            zoomOutBtn.title = 'Exit zoom mode and show all nodes';

            bottomControls.appendChild(zoomOutBtn);
        }
    }

    /**
     * Remove zoom-out button from tree controls
     */
    removeZoomOutButton() {
        const existingButton = document.getElementById('zoom-out-btn');
        if (existingButton) {
            existingButton.remove();
        }
    }

    /**
     * Get the appropriate icon for a node type
     * @param {string} type - The node type
     * @param {Object} node - The node object to check for specific conditions
     * @returns {string} - The emoji icon
     */
    getNodeIcon(type, node = null) {
        // Check if this is a PowerShell script execution
        if (node) {
            const nodeTitle = this.getNodeTitle(node);
            const nodeSubtitle = this.getNodeSubtitle(node);
            const isPowerShellScript = (nodeTitle && nodeTitle.toLowerCase().includes('powershell.exe executed a script')) ||
                (nodeSubtitle && nodeSubtitle.toLowerCase().includes('powershell.exe executed a script'));
            if (isPowerShellScript) {
                return '📜'; // Script icon for PowerShell executions
            }
        }

        const icons = {
            'process': '⚙️',
            'file': '📄',
            'account': '👤',
            'network': '🌐',
            'registry': '📋',
            'url': '🔗',
            'ip': '🌐',
            'domain': '🌍',
            'other': '📦'
        };
        return icons[type] || '📦';
    }

    /**
     * Extract the main title from a node
     * @param {Object} node - The node
     * @returns {string} - The node title
     */
    getNodeTitle(node) {
        if (node.title) {
            const parts = [];
            if (node.title.prefix && node.title.prefix.trim()) {
                parts.push(node.title.prefix);
            }
            if (node.title.main) {
                parts.push(node.title.main);
            }
            return parts.join(' ') || 'Unknown';
        }

        if (node.entity) {
            if (node.entity.ImageFile && node.entity.ImageFile.FileName) {
                return node.entity.ImageFile.FileName;
            }
            if (node.entity.User && node.entity.User.UserName) {
                return `${node.entity.User.DomainName || ''}\\${node.entity.User.UserName}`;
            }
        }

        return 'Unknown';
    }

    /**
     * Extract the subtitle from a node
     * @param {Object} node - The node
     * @returns {string|null} - The node subtitle
     */
    getNodeSubtitle(node) {
        return node.title && node.title.intro ? node.title.intro : null;
    }

    /**
     * Extract the command line from a node's entity
     * @param {Object} node - The node
     * @returns {string|null} - The command line or null
     */
    getNodeCommandLine(node) {
        // First check for command line in entity
        if (node.entity && node.entity.Commandline) {
            return node.entity.Commandline;
        }

        // Also check for WMI Query in node details
        if (node.details && Array.isArray(node.details)) {
            const wmiQuery = node.details.find(detail =>
                detail.key && detail.key.toLowerCase().includes('wmi query') && detail.value
            );
            if (wmiQuery) {
                return `WMI: ${wmiQuery.value}`;
            }
        }

        return null;
    }

    /**
     * Format timestamp for display
     * @param {string} timeString - The timestamp string
     * @returns {string|null} - Formatted time or null
     */
    formatTime(timeString) {
        if (!timeString) return null;

        try {
            const date = new Date(timeString);
            return date.toLocaleString('en-US', {
                year: 'numeric',
                month: 'short',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
        } catch (error) {
            console.warn('Invalid date format:', timeString);
            return timeString;
        }
    }

    /**
     * Escape HTML characters to prevent XSS
     * @param {string} text - The text to escape
     * @returns {string} - Escaped text
     */
    escapeHtml(text) {
        if (typeof text !== 'string') return String(text);

        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * Render detailed information for a node
     * @param {Object} node - The node to render details for
     * @returns {string} - HTML string for the details
     */
    renderNodeDetails(node) {
        let html = '';

        // Entity information
        if (node.entity) {
            html += this.renderEntityDetails(node.entity);
        }

        // Node details
        if (node.details && Array.isArray(node.details) && node.details.length > 0) {
            html += this.renderDetailsSection('Node Details', node.details);
        }

        // Additional details
        if (node.additionalDetails && Array.isArray(node.additionalDetails)) {
            node.additionalDetails.forEach(section => {
                if (section.title && section.details) {
                    const title = section.title.main || 'Additional Details';
                    html += this.renderDetailsSection(title, section.details);
                }
            });
        }

        return html || '<div class="details-section"><div class="details-title">No additional details available</div></div>';
    }

    /**
     * Render entity-specific details
     * @param {Object} entity - The entity object
     * @returns {string} - HTML string for entity details
     */
    renderEntityDetails(entity) {
        // Check if entity has meaningful data first
        if (!this.entityHasMeaningfulData(entity)) {
            return '';
        }

        let html = '<div class="details-section"><div class="details-title">Entity Information</div>';

        // Image file information
        if (entity.ImageFile) {
            const img = entity.ImageFile;
            html += this.renderDetailItem('File Path', img.FullPath);
            html += this.renderDetailItem('File Size', img.Size ? `${(img.Size / 1024).toFixed(2)} KB` : null);
            html += this.renderDetailItem('SHA256', img.Sha256);
            html += this.renderDetailItem('SHA1', img.Sha1);
            html += this.renderDetailItem('MD5', img.Md5);
            html += this.renderDetailItem('Creation Time', this.formatTime(img.CreationTime));
        }

        // User information
        if (entity.User) {
            const user = entity.User;
            html += this.renderDetailItem('Domain', user.DomainName);
            html += this.renderDetailItem('Username', user.UserName);
            html += this.renderDetailItem('SID', user.Sid);
        }

        // Process information
        if (entity.ProcessId) {
            html += this.renderDetailItem('Process ID', entity.ProcessId);
            html += this.renderDetailItem('Command Line', entity.Commandline, 'script');
            html += this.renderDetailItem('Parent Process ID', entity.CreatingProcessId);
            html += this.renderDetailItem('Parent Process', entity.CreatingProcessName);
            html += this.renderDetailItem('Creation Time', this.formatTime(entity.CreationTime));
            html += this.renderDetailItem('Integrity Level', entity.IntegrityLevel);
            html += this.renderDetailItem('Token Elevation', entity.TokenElevation);
        }

        html += '</div>';
        return html;
    }

    /**
     * Render a section of details
     * @param {string} title - The section title
     * @param {Array} details - Array of detail objects
     * @returns {string} - HTML string for the section
     */
    renderDetailsSection(title, details) {
        let html = `<div class="details-section"><div class="details-title">${this.escapeHtml(title)}</div>`;

        details.forEach(detail => {
            if (detail.key && detail.value !== undefined && detail.value !== null) {
                const valueType = detail.valueType === 'script' ? 'script' : null;
                let value = detail.value;

                // Format dates
                if (detail.valueType === 'date') {
                    value = this.formatTime(value) || value;
                }

                html += this.renderDetailItem(detail.key, value, valueType);
            }
        });

        html += '</div>';
        return html;
    }

    /**
     * Render a single detail item
     * @param {string} key - The detail key
     * @param {string} value - The detail value
     * @param {string} valueType - The value type for styling
     * @returns {string} - HTML string for the detail item
     */
    renderDetailItem(key, value, valueType = null) {
        if (value === undefined || value === null || value === '') {
            return '';
        }

        // Treat WMI Query fields the same as command lines (script styling)
        const isWmiQuery = key.toLowerCase().includes('wmi query');
        const isCommandLine = key.toLowerCase().includes('command line');
        const isScript = valueType === 'script' || isWmiQuery;
        const valueClass = isScript ? 'detail-value script' : 'detail-value';

        // Handle script content (especially "Content" fields) with proper newline display
        let displayValue;
        if (isScript && key.toLowerCase().includes('content')) {
            // For script content, convert escaped newlines to actual line breaks and preserve formatting
            displayValue = this.escapeHtml(this.unescapeForwardSlashes(String(value)))
                .replace(/\\r\\n/g, '<br>')
                .replace(/\\n/g, '<br>')
                .replace(/\\r/g, '<br>');
        } else if (isCommandLine || isScript || isWmiQuery) {
            // For command lines and script values, unescape forward slashes
            displayValue = this.escapeHtml(this.unescapeForwardSlashes(String(value)));
        } else {
            displayValue = this.escapeHtml(String(value));
        }

        return `
            <div class="detail-item">
                <div class="detail-key">${this.escapeHtml(key)}:</div>
                <div class="${valueClass}">${displayValue}</div>
            </div>
        `;
    }

    /**
     * Minimize the upload section after successful file processing
     */
    minimizeUploadSection() {
        const uploadArea = document.getElementById('upload-area');
        const uploadText = uploadArea.querySelector('.upload-text');

        if (uploadArea && uploadText) {
            uploadArea.classList.add('minimized');

            // Update text to show current file status
            uploadText.textContent = 'File loaded successfully - Click to upload a different file';
        }
    }

    /**
     * Restore the upload section to full size
     */
    restoreUploadSection() {
        const uploadArea = document.getElementById('upload-area');
        const uploadText = uploadArea.querySelector('.upload-text');

        if (uploadArea && uploadText) {
            uploadArea.classList.remove('minimized');
            uploadText.textContent = 'Drop your JSOC file here or click to browse';
        }
    }

    /**
     * Toggle node details panel
     * @param {string} nodeId - The node ID
     */
    toggleNodeDetails(nodeId) {
        const nodeElement = document.querySelector(`[data-node-id="${nodeId}"]`);
        if (!nodeElement) return;

        const detailsPanel = nodeElement.querySelector('.details-panel');
        if (!detailsPanel) return;

        if (detailsPanel.style.display === 'none' || detailsPanel.style.display === '') {
            detailsPanel.style.display = 'block';
        } else {
            detailsPanel.style.display = 'none';
        }
    }

    /**
     * Show loading state
     */
    showLoading() {
        const processTree = document.getElementById('process-tree');
        const treeContent = document.getElementById('tree-content');

        if (processTree) processTree.style.display = 'block';
        if (treeContent) treeContent.innerHTML = '<div class="loading">Loading and parsing file...</div>';
    }

    /**
     * Download the redacted/anonymized JSON data
     */
    downloadRedactedJson() {
        if (!this.originalData) {
            console.error('No data available for download');
            return;
        }

        try {
            // Create a copy of the data to avoid modifying the original
            let dataToDownload = JSON.parse(JSON.stringify(this.originalData));

            // Apply anonymization if enabled
            if (this.isAnonymized) {
                dataToDownload = this.createAnonymizedData(dataToDownload);
            }

            // Create filename with timestamp
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
            const suffix = this.isAnonymized ? '_anonymized' : '';
            const filename = `xdr_story_data${suffix}_${timestamp}.json`;

            // Create and trigger download
            const dataStr = JSON.stringify(dataToDownload, null, 2);
            const dataBlob = new Blob([dataStr], { type: 'application/json' });

            const downloadLink = document.createElement('a');
            downloadLink.href = URL.createObjectURL(dataBlob);
            downloadLink.download = filename;
            downloadLink.style.display = 'none';

            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);

            // Clean up the object URL
            URL.revokeObjectURL(downloadLink.href);

            console.log(`Downloaded ${this.isAnonymized ? 'anonymized' : 'redacted'} JSON as ${filename}`);

        } catch (error) {
            console.error('Error downloading JSON:', error);
            this.showError('Failed to download JSON file');
        }
    }

    /**
     * Capture a screenshot of the current process tree
     */
    async captureScreenshot() {
        if (!this.data || !this.data.items) {
            this.showError('No data loaded. Please upload a file first.');
            return;
        }

        try {
            console.log('Starting screenshot capture...');

            // Show loading state
            const screenshotBtn = document.getElementById('screenshot-btn');
            const originalText = screenshotBtn.innerHTML;
            screenshotBtn.innerHTML = '📸 Capturing...';
            screenshotBtn.disabled = true;

            // Get the tree-container element
            const treeContainer = document.getElementById('tree-content');

            if (!treeContainer) {
                throw new Error('Tree container not found');
            }

            // Clone the tree container
            const treeClone = treeContainer.cloneNode(true);

            // Create a temporary container for capture
            const captureContainer = document.createElement('div');
            captureContainer.style.cssText = `
                position: absolute;
                top: -10000px;
                left: -10000px;
                background: #0d1117;
                padding: 25px;
                border-radius: 12px;
                font-family: 'SF Mono', 'Monaco', 'Inconsolata', 'Roboto Mono', 'Courier New', monospace;
                font-size: 14px;
                line-height: 1.5;
                color: #f0f6fc;
                border: 1px solid #30363d;
                box-sizing: border-box;
            `;

            // Remove max-height and overflow constraints from the clone to show full content
            treeClone.style.maxHeight = 'none';
            treeClone.style.overflow = 'visible';
            treeClone.style.height = 'auto';

            // Append the clone to our capture container
            captureContainer.appendChild(treeClone);

            // Temporarily add to body for rendering
            document.body.appendChild(captureContainer);

            // Wait a brief moment for styles to apply
            await new Promise(resolve => setTimeout(resolve, 100));

            // Capture with html2canvas
            const canvas = await html2canvas(captureContainer, {
                backgroundColor: '#0d1117',
                scale: 2, // Higher resolution
                useCORS: true,
                allowTaint: true,
                logging: false,
                width: captureContainer.scrollWidth,
                height: captureContainer.scrollHeight,
                scrollX: 0,
                scrollY: 0,
                onclone: (clonedDoc) => {
                    // Ensure the cloned container shows full content
                    const clonedContainer = clonedDoc.querySelector('div');
                    if (clonedContainer) {
                        clonedContainer.style.maxHeight = 'none';
                        clonedContainer.style.overflow = 'visible';
                        clonedContainer.style.height = 'auto';
                    }
                }
            });

            // Clean up temporary container
            document.body.removeChild(captureContainer);

            // Create download link
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
            const modeText = this.isZoomedMode ? '_zoomed' : '';
            const anonymizedText = this.isAnonymized ? '_anonymized' : '';
            const filename = `xdr_process_tree${modeText}${anonymizedText}_${timestamp}.png`;

            // Convert canvas to blob and download
            canvas.toBlob((blob) => {
                const downloadLink = document.createElement('a');
                downloadLink.href = URL.createObjectURL(blob);
                downloadLink.download = filename;
                downloadLink.style.display = 'none';

                document.body.appendChild(downloadLink);
                downloadLink.click();
                document.body.removeChild(downloadLink);

                // Clean up
                URL.revokeObjectURL(downloadLink.href);

                console.log(`Screenshot saved as ${filename}`);

                // Show success message temporarily
                screenshotBtn.innerHTML = '✅ Captured!';
                setTimeout(() => {
                    screenshotBtn.innerHTML = originalText;
                    screenshotBtn.disabled = false;
                }, 2000);
            }, 'image/png', 0.95);

        } catch (error) {
            console.error('Error capturing screenshot:', error);

            // Reset button
            const screenshotBtn = document.getElementById('screenshot-btn');
            if (screenshotBtn) {
                screenshotBtn.innerHTML = '📸 Screenshot';
                screenshotBtn.disabled = false;
            }

            // Show error message
            this.showError('Failed to capture screenshot. Please try again.');
        }
    }

    /**
     * Show error message
     * @param {string} message - The error message
     */
    showError(message) {
        const processTree = document.getElementById('process-tree');
        const treeContent = document.getElementById('tree-content');
        const investigationInfo = document.getElementById('investigation-info');
        const anonymizeCheckbox = document.getElementById('anonymize-checkbox');

        this.restoreUploadSection();

        if (anonymizeCheckbox) anonymizeCheckbox.checked = false;
        this.isAnonymized = false;

        if (investigationInfo) investigationInfo.style.display = 'none';

        const downloadBtn = document.getElementById('download-json-btn');
        if (downloadBtn) downloadBtn.style.display = 'none';

        if (processTree) processTree.style.display = 'block';
        if (treeContent) treeContent.innerHTML = `<div class="error">${this.escapeHtml(message)}</div>`;

        console.error('XDR Visualizer Error:', message);
    }

    /**
     * Extract all command lines from the data, sorted by timestamp
     */
    extractCommandLines() {
        if (!this.data || !this.data.items) {
            this.showError('No data loaded. Please upload a file first.');
            return;
        }

        console.log('Extracting command lines...');

        const commandLines = [];

        if (this.isZoomedMode && this.zoomedNodeId) {
            console.log('Zoom mode active - extracting from visible nodes only');

            const visibleNodes = document.querySelectorAll('.tree-node');
            const visibleTreeNodes = Array.from(visibleNodes).filter(node => node.offsetParent !== null);

            console.log(`Found ${visibleTreeNodes.length} visible nodes in zoom mode`);

            visibleTreeNodes.forEach(nodeElement => {
                const nodeId = nodeElement.dataset.nodeId;
                const item = this.findItemById(nodeId, this.data.items);

                if (item) {
                    const commandLine = this.getCommandLineFromItem(item);
                    if (commandLine && commandLine.trim() !== '') {
                        const timestamp = this.getTimestampFromItem(item);
                        const processName = this.getProcessNameFromItem(item);
                        const userInfo = this.getUserInfoFromItem(item);

                        commandLines.push({
                            timestamp: timestamp,
                            processName: processName,
                            userInfo: userInfo,
                            commandLine: commandLine.trim()
                        });
                    }
                }
            });
        } else {
            // Normal mode: extract from all items recursively
            console.log('Normal mode - extracting from all nodes');

            // Function to recursively extract command lines from items
            const extractFromItems = (items, depth = 0) => {
                if (!Array.isArray(items)) return;

                items.forEach(item => {
                    // Check if this item has a command line
                    const commandLine = this.getCommandLineFromItem(item);
                    if (commandLine && commandLine.trim() !== '') {
                        const timestamp = this.getTimestampFromItem(item);
                        const processName = this.getProcessNameFromItem(item);
                        const userInfo = this.getUserInfoFromItem(item);

                        commandLines.push({
                            timestamp: timestamp,
                            processName: processName,
                            userInfo: userInfo,
                            commandLine: commandLine.trim()
                        });
                    }

                    // Recursively check children
                    if (item.children && Array.isArray(item.children)) {
                        extractFromItems(item.children, depth + 1);
                    }
                    if (item.nestedItems && Array.isArray(item.nestedItems)) {
                        extractFromItems(item.nestedItems, depth + 1);
                    }
                });
            };

            // Extract from root items
            extractFromItems(this.data.items);
        }

        // Sort by timestamp (ascending)
        commandLines.sort((a, b) => {
            if (!a.timestamp && !b.timestamp) return 0;
            if (!a.timestamp) return 1;
            if (!b.timestamp) return -1;
            return new Date(a.timestamp) - new Date(b.timestamp);
        });

        console.log(`Found ${commandLines.length} command lines`);

        // Format the output
        let output = '';
        commandLines.forEach((cmd, index) => {
            const timeStr = cmd.timestamp ? new Date(cmd.timestamp).toLocaleString() : 'No timestamp';
            // Unescape forward slashes in command line for better readability
            const unescapedCommandLine = this.unescapeForwardSlashes(cmd.commandLine);
            output += `# ${timeStr} - ${cmd.processName} - User: ${cmd.userInfo}\n`;
            output += `${unescapedCommandLine}\n\n`;
        });

        if (output === '') {
            const modeText = this.isZoomedMode ? 'visible (zoomed)' : 'loaded';
            output = `# No command lines found in the ${modeText} data\n# This might indicate that the data doesn't contain process creation events with command lines`;
        }

        // Update the title and display in shared textarea
        const titleElement = document.getElementById('analysis-output-title');
        if (titleElement) {
            const modeText = this.isZoomedMode ? ' (zoomed view)' : '';
            titleElement.textContent = `Command Lines${modeText} (sorted by timestamp)`;
        }

        const textarea = document.getElementById('analysis-output');
        if (textarea) {
            textarea.value = output;
        }

        // Show the copy button
        const copyBtn = document.getElementById('copy-analysis-btn');
        if (copyBtn) {
            copyBtn.style.display = 'inline-block';
        }

        // Show the analysis tools section
        const analysisSection = document.getElementById('analysis-tools');
        if (analysisSection) {
            analysisSection.style.display = 'block';
        }
    }

    /**
     * Extract all PowerShell scripts from the data, sorted by timestamp
     */
    extractPowerShellScripts() {
        if (!this.data || !this.data.items) {
            this.showError('No data loaded. Please upload a file first.');
            return;
        }

        console.log('Extracting PowerShell scripts...');

        const scripts = [];

        if (this.isZoomedMode && this.zoomedNodeId) {
            console.log('Zoom mode active - extracting PowerShell scripts from visible nodes only');

            const visibleNodes = document.querySelectorAll('.tree-node');
            const visibleTreeNodes = Array.from(visibleNodes).filter(node => node.offsetParent !== null);

            console.log(`Found ${visibleTreeNodes.length} visible nodes in zoom mode`);

            visibleTreeNodes.forEach(nodeElement => {
                const nodeId = nodeElement.dataset.nodeId;
                const item = this.findItemById(nodeId, this.data.items);

                if (item) {
                    const scriptContent = this.getPowerShellScriptFromItem(item);
                    if (scriptContent && scriptContent.trim() !== '') {
                        const timestamp = this.getTimestampFromItem(item);
                        const processName = this.getProcessNameFromItem(item);
                        const userInfo = this.getUserInfoFromItem(item);

                        scripts.push({
                            timestamp: timestamp,
                            processName: processName,
                            userInfo: userInfo,
                            scriptContent: scriptContent.trim()
                        });
                    }
                }
            });
        } else {
            // Normal mode: extract from all items recursively
            console.log('Normal mode - extracting PowerShell scripts from all nodes');

            // Function to recursively extract PowerShell scripts from items
            const extractFromItems = (items, depth = 0) => {
                if (!Array.isArray(items)) return;

                items.forEach(item => {
                    // Check if this item has a PowerShell script
                    const scriptContent = this.getPowerShellScriptFromItem(item);
                    if (scriptContent && scriptContent.trim() !== '') {
                        const timestamp = this.getTimestampFromItem(item);
                        const processName = this.getProcessNameFromItem(item);
                        const userInfo = this.getUserInfoFromItem(item);

                        scripts.push({
                            timestamp: timestamp,
                            processName: processName,
                            userInfo: userInfo,
                            scriptContent: scriptContent.trim()
                        });
                    }

                    // Recursively check children
                    if (item.children && Array.isArray(item.children)) {
                        extractFromItems(item.children, depth + 1);
                    }
                    if (item.nestedItems && Array.isArray(item.nestedItems)) {
                        extractFromItems(item.nestedItems, depth + 1);
                    }
                });
            };

            // Extract from root items
            extractFromItems(this.data.items);
        }

        // Sort by timestamp (ascending)
        scripts.sort((a, b) => {
            if (!a.timestamp && !b.timestamp) return 0;
            if (!a.timestamp) return 1;
            if (!b.timestamp) return -1;
            return new Date(a.timestamp) - new Date(b.timestamp);
        });

        console.log(`Found ${scripts.length} PowerShell scripts`);

        // Format the output
        let output = '';
        scripts.forEach((script, index) => {
            const timeStr = script.timestamp ? new Date(script.timestamp).toLocaleString() : 'No timestamp';
            // Unescape forward slashes and other escaped characters for better readability
            const unescapedScript = this.unescapeScriptContent(script.scriptContent);
            output += `# ${timeStr} - ${script.processName} - User: ${script.userInfo}\n`;
            output += `${unescapedScript}\n\n`;
            output += `# ${'='.repeat(80)}\n\n`; // Add separator between scripts
        });

        if (output === '') {
            const modeText = this.isZoomedMode ? 'visible (zoomed)' : 'loaded';
            output = `# No PowerShell scripts found in the ${modeText} data\n# This might indicate that the data doesn't contain "powershell.exe executed a script" events`;
        }

        // Update the title and display in shared textarea
        const titleElement = document.getElementById('analysis-output-title');
        if (titleElement) {
            const modeText = this.isZoomedMode ? ' (zoomed view)' : '';
            titleElement.textContent = `PowerShell Scripts${modeText} (sorted by timestamp)`;
        }

        const textarea = document.getElementById('analysis-output');
        if (textarea) {
            textarea.value = output;
        }

        // Show the copy button
        const copyBtn = document.getElementById('copy-analysis-btn');
        if (copyBtn) {
            copyBtn.style.display = 'inline-block';
        }

        // Show the analysis tools section
        const analysisSection = document.getElementById('analysis-tools');
        if (analysisSection) {
            analysisSection.style.display = 'block';
        }
    }

    /**
     * Copy analysis results to clipboard
     */
    copyAnalysisResults() {
        const textarea = document.getElementById('analysis-output');
        if (textarea && textarea.value) {
            navigator.clipboard.writeText(textarea.value).then(() => {
                // Temporarily change button text to show success
                const copyBtn = document.getElementById('copy-analysis-btn');
                if (copyBtn) {
                    const originalText = copyBtn.innerHTML;
                    copyBtn.innerHTML = '✅ Copied!';
                    setTimeout(() => {
                        copyBtn.innerHTML = originalText;
                    }, 2000);
                }
            }).catch(err => {
                console.error('Failed to copy to clipboard:', err);
                // Fallback: select the text
                textarea.select();
                textarea.setSelectionRange(0, 99999); // For mobile devices
            });
        }
    }

    /**
     * Extract command line from an item using the same logic as getNodeCommandLine
     * @param {Object} item - The item to extract command line from
     * @returns {string|null} - The command line or null
     */
    getCommandLineFromItem(item) {
        // First check for command line in entity
        if (item.entity && item.entity.Commandline) {
            return item.entity.Commandline;
        }

        // Also check for WMI Query in item details
        if (item.details && Array.isArray(item.details)) {
            const wmiQuery = item.details.find(detail =>
                detail.key && detail.key.toLowerCase().includes('wmi query') && detail.value
            );
            if (wmiQuery) {
                return `WMI: ${wmiQuery.value}`;
            }
        }

        return null;
    }

    /**
     * Extract timestamp from an item
     * @param {Object} item - The item to extract timestamp from
     * @returns {string|null} - The timestamp or null
     */
    getTimestampFromItem(item) {
        // Check various timestamp fields
        if (item.time) return item.time;
        if (item.entity && item.entity.CreationTime) return item.entity.CreationTime;
        if (item.processCreationTime) return item.processCreationTime;
        if (item.timestamp) return item.timestamp;
        return null;
    }

    /**
     * Extract process name from an item using similar logic as getNodeTitle
     * @param {Object} item - The item to extract process name from
     * @returns {string} - The process name
     */
    getProcessNameFromItem(item) {
        // Check title first
        if (item.title && item.title.main) {
            return item.title.main;
        }

        // Check entity
        if (item.entity) {
            if (item.entity.ImageFile && item.entity.ImageFile.FileName) {
                return item.entity.ImageFile.FileName;
            }
            if (item.entity.User && item.entity.User.UserName) {
                return `${item.entity.User.DomainName || ''}\\${item.entity.User.UserName}`;
            }
        }

        // Fallback checks
        if (item.fileName) return item.fileName;
        if (item.processName) return item.processName;
        if (item.name) return item.name;

        return 'Unknown Process';
    }

    /**
     * Extract user information from an item
     * @param {Object} item - The item to extract user info from
     * @returns {string} - The user information
     */
    getUserInfoFromItem(item) {
        // Check entity first
        if (item.entity && item.entity.User) {
            const user = item.entity.User;
            if (user.UserName) {
                if (user.DomainName) {
                    return `${user.DomainName}\\${user.UserName}`;
                }
                return user.UserName;
            }
            if (user.Sid) {
                return `SID: ${user.Sid}`;
            }
        }

        // Check direct properties
        if (item.accountName) {
            if (item.accountDomain) {
                return `${item.accountDomain}\\${item.accountName}`;
            }
            return item.accountName;
        }

        if (item.initiatingProcessAccountName) {
            if (item.initiatingProcessAccountDomain) {
                return `${item.initiatingProcessAccountDomain}\\${item.initiatingProcessAccountName}`;
            }
            return item.initiatingProcessAccountName;
        }

        if (item.accountUpn) {
            return item.accountUpn;
        }

        if (item.accountSid) {
            return `SID: ${item.accountSid}`;
        }

        return 'Unknown User';
    }

    /**
     * Unescape forward slashes for better readability in command lines
     * @param {string} str - The string with escaped forward slashes
     * @returns {string} - The string with unescaped forward slashes
     */
    unescapeForwardSlashes(str) {
        if (!str || typeof str !== 'string') return str;
        return str.replace(/\\\//g, '/');
    }

    /**
     * Find an item by its node ID in the data structure
     * @param {string} nodeId - The node ID to search for
     * @param {Array} items - The items array to search in
     * @returns {Object|null} - The found item or null
     */
    findItemById(nodeId, items) {
        if (!Array.isArray(items)) return null;

        for (const item of items) {
            // Check if this item matches the ID
            const itemId = this.getNodeId(item);
            if (itemId === nodeId) {
                return item;
            }

            // Search in children
            if (item.children && Array.isArray(item.children)) {
                const found = this.findItemById(nodeId, item.children);
                if (found) return found;
            }

            // Search in nested items
            if (item.nestedItems && Array.isArray(item.nestedItems)) {
                const found = this.findItemById(nodeId, item.nestedItems);
                if (found) return found;
            }
        }

        return null;
    }

    /**
     * Extract PowerShell script content from an item
     * @param {Object} item - The item to extract PowerShell script from
     * @returns {string|null} - The PowerShell script content or null
     */
    getPowerShellScriptFromItem(item) {
        // Check if this is a PowerShell script execution event
        const nodeTitle = this.getNodeTitle(item);
        const nodeSubtitle = this.getNodeSubtitle(item);

        const isPowerShellScript = (nodeTitle && nodeTitle.toLowerCase().includes('powershell.exe executed a script')) ||
            (nodeSubtitle && nodeSubtitle.toLowerCase().includes('powershell.exe executed a script'));

        if (!isPowerShellScript) {
            return null;
        }

        // Look for Content in the details
        if (item.details && Array.isArray(item.details)) {
            const contentDetail = item.details.find(detail =>
                detail.key && detail.key.toLowerCase() === 'content' && detail.value
            );
            if (contentDetail) {
                return contentDetail.value;
            }
        }

        // Also check in additionalDetails
        if (item.additionalDetails && Array.isArray(item.additionalDetails)) {
            for (const section of item.additionalDetails) {
                if (section.details && Array.isArray(section.details)) {
                    const contentDetail = section.details.find(detail =>
                        detail.key && detail.key.toLowerCase() === 'content' && detail.value
                    );
                    if (contentDetail) {
                        return contentDetail.value;
                    }
                }
            }
        }

        return null;
    }

    /**
     * Unescape script content for better readability
     * @param {string} str - The script content with escaped characters
     * @returns {string} - The script content with unescaped characters
     */
    unescapeScriptContent(str) {
        if (!str || typeof str !== 'string') return str;

        return str
            .replace(/\\\//g, '/') // Unescape forward slashes
            .replace(/\\r\\n/g, '\n') // Convert Windows line breaks
            .replace(/\\n/g, '\n') // Convert Unix line breaks  
            .replace(/\\r/g, '\n') // Convert Mac line breaks
            .replace(/\\t/g, '\t') // Convert tabs
            .replace(/\\"/g, '"') // Unescape double quotes
            .replace(/\\\\/g, '\\'); // Unescape backslashes (do this last)
    }

    /**
     * Change the application theme
     * @param {string} themeName - The theme to apply ('cyberpunk' or 'professional')
     */
    changeTheme(themeName) {
        const body = document.body;

        // Remove existing theme classes
        body.classList.remove('theme-cyberpunk', 'theme-professional');

        // Add new theme class
        body.classList.add(`theme-${themeName}`);

        localStorage.setItem('xdr-theme', themeName);

        console.log(`Theme changed to: ${themeName}`);
    }

    /**
     * Initialize theme from localStorage or default
     */
    initializeTheme() {
        const savedTheme = localStorage.getItem('xdr-theme') || 'cyberpunk';
        const themeSelect = document.getElementById('header-theme-select');

        if (themeSelect) {
            themeSelect.value = savedTheme;
        }

        this.changeTheme(savedTheme);
    }
}

// Global instance
let xdrVisualizer;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    xdrVisualizer = new XDRTreeVisualizer();
    console.log('XDR Tree Visualizer initialized');
});
