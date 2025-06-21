(function() {
    'use strict';

    class PoemCompiler {
        constructor() {
            this.poems = [];
            this.selectedFiles = [];
            this.draggedIndex = null;
            this.isProcessing = false;
            this.notificationTimeout = null;
            console.log('PoemCompiler initialized.');
            this.initializeEventListeners();
            this.updateDisplay();
        }

        /**
         * Initializes all event listeners for the UI elements.
         */
        initializeEventListeners() {
            console.log('Initializing event listeners...');
            const wordFiles = document.getElementById('wordFiles');
            const processBtn = document.getElementById('processBtn');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');
            const fileLabel = document.getElementById('fileLabel');

            if (!wordFiles || !processBtn || !downloadBtn || !clearBtn || !fileLabel) {
                console.error('Required DOM elements not found. Please ensure all IDs are correct in the HTML.');
                this.showNotification('Application setup error: Some UI elements are missing. Please check the HTML.', 'error', 0);
                return;
            }

            // File input change event
            wordFiles.addEventListener('change', (e) => {
                console.log('File input change event triggered.');
                this.handleFileSelect(e);
            });

            // Process button click event
            processBtn.addEventListener('click', () => {
                console.log('Process button clicked. current this.isProcessing:', this.isProcessing);
                if (!this.isProcessing) {
                    this.processDocuments();
                } else {
                    this.showNotification('Already processing documents. Please wait.', 'info');
                }
            });

            // Download button click event
            downloadBtn.addEventListener('click', () => {
                console.log('Download button clicked.');
                this.downloadCombinedDocument();
            });

            // Clear button click event
            clearBtn.addEventListener('click', () => {
                console.log('Clear button clicked.');
                this.clearAllPoems();
            });

            // --- Drag and drop functionality for the file label ---
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, (e) => this.preventDefaults(e), false);
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#3b82f6';
                    fileLabel.style.backgroundColor = '#eff6ff';
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#d1d5db';
                    fileLabel.style.backgroundColor = '';
                }, false);
            });

            fileLabel.addEventListener('drop', (e) => {
                console.log('File dropped.');
                const files = Array.from(e.dataTransfer.files).filter(file =>
                    file.name.toLowerCase().endsWith('.docx')
                );
                if (files.length > 0) {
                    const dt = new DataTransfer();
                    files.forEach(file => dt.items.add(file));
                    wordFiles.files = dt.files;
                    const event = new Event('change', { bubbles: true });
                    wordFiles.dispatchEvent(event); // Trigger change event programmatically
                    console.log('Dropped .docx files, dispatched change event.');
                } else if (e.dataTransfer.files.length > 0) {
                    this.showNotification('Please upload only .docx files', 'warning');
                    console.warn('Dropped files but none were .docx.');
                }
            }, false);
            console.log('Event listeners initialized successfully.');
        }

        /**
         * Prevents default event behaviors (e.g., opening dropped files).
         * @param {Event} e - The event object.
         */
        preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        /**
         * Handles file selection from the input or drag-and-drop.
         * Validates file types and updates the UI accordingly.
         * @param {Event} event - The change event from the file input.
         */
        handleFileSelect(event) {
            const files = Array.from(event.target.files);
            console.log('Files selected in handleFileSelect:', files);

            const fileLabel = document.getElementById('fileLabel');
            const processBtn = document.getElementById('processBtn');

            const validFiles = files.filter(file => file.name.toLowerCase().endsWith('.docx'));
            const invalidFiles = files.filter(file => !file.name.toLowerCase().endsWith('.docx'));

            if (invalidFiles.length > 0) {
                this.showNotification(`${invalidFiles.length} invalid file(s) ignored. Only .docx files are supported.`, 'warning');
                console.warn(`${invalidFiles.length} invalid file(s) ignored.`);
            }

            if (validFiles.length > 0) {
                this.selectedFiles = validFiles;
                console.log('Valid files assigned to this.selectedFiles:', this.selectedFiles);
                const fileNames = validFiles.length > 3
                    ? `${validFiles.slice(0, 3).map(f => f.name).join(', ')} and ${validFiles.length - 3} more...`
                    : validFiles.map(f => f.name).join(', ');

                fileLabel.innerHTML = `
                    <span>📄</span>
                    <span>Selected: ${validFiles.length} document${validFiles.length > 1 ? 's' : ''}</span>
                    <small>${this.escapeHtml(fileNames)}</small>
                `;
                fileLabel.classList.add('has-files');
                processBtn.disabled = false;
                this.announceToScreenReader('process-status', `${validFiles.length} documents selected, ready to process`);
                console.log(`UI updated: ${validFiles.length} documents selected, process button enabled.`);
            } else {
                this.selectedFiles = [];
                fileLabel.innerHTML = `
                    <span>📄</span>
                    <span>Click here or drag Word documents to upload</span>
                    <small>Multiple files supported</small>
                `;
                fileLabel.classList.remove('has-files');
                processBtn.disabled = true;
                this.announceToScreenReader('process-status', 'No valid documents selected');
                console.log('No valid files selected. UI reset.');
            }
        }

        /**
         * Processes the selected Word documents to extract poems.
         * Displays progress and notifications.
         */
        async processDocuments() {
            if (this.selectedFiles.length === 0) {
                this.showNotification('Please select Word documents first!', 'warning');
                console.warn('ProcessDocuments called with no selected files.');
                return;
            }

            if (this.isProcessing) {
                console.log('Attempted to process while already processing.');
                return;
            }

            this.isProcessing = true;
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');
            const progressBar = document.getElementById('progressBar');

            console.log('Starting processDocuments. Updating UI...');
            processBtn.disabled = true;
            processBtn.textContent = 'Processing...';
            progressContainer.style.display = 'block';
            progressBar.style.width = '0%';
            progressBar.setAttribute('aria-valuenow', '0');

            this.announceToScreenReader('process-status', 'Processing documents...');

            try {
                let processedPoemCount = 0;
                let skippedCount = 0;
                const totalFiles = this.selectedFiles.length;
                console.log(`Processing ${totalFiles} selected files.`);
                const errors = [];

                for (let i = 0; i < this.selectedFiles.length; i++) {
                    const file = this.selectedFiles[i];
                    console.log(`Processing file ${i + 1}/${totalFiles}: ${file.name}`);

                    try {
                        const poemsFromFile = await this.extractPoemsFromDocument(file);
                        console.log(`Extracted ${poemsFromFile ? poemsFromFile.length : 0} potential poems from ${file.name}`);
                        if (poemsFromFile && poemsFromFile.length > 0) {
                            for (const poemData of poemsFromFile) {
                                if (poemData && poemData.content && poemData.content.trim().length > 0) {
                                    const isDuplicate = this.poems.some(existing =>
                                        existing.title.toLowerCase() === poemData.title.toLowerCase() &&
                                        (existing.content.trim().length > 50 && existing.content.trim() === poemData.content.trim())
                                    );

                                    if (!isDuplicate) {
                                        this.poems.push(poemData);
                                        processedPoemCount++;
                                        console.log(`Added new poem: "${poemData.title}" from "${file.name}"`);
                                    } else {
                                        skippedCount++;
                                        console.warn(`Duplicate poem detected and skipped: "${poemData.title || 'Untitled'}" from "${file.name}"`);
                                    }
                                } else {
                                    console.warn(`Poem data from ${file.name} was empty or invalid.`);
                                }
                            }
                        } else {
                            errors.push(`${file.name}: No valid poems found`);
                            console.warn(`No valid poems found in ${file.name}.`);
                        }
                    } catch (error) {
                        console.error(`Error processing ${file.name}:`, error);
                        errors.push(`${file.name}: ${error.message}`);
                    }

                    const progress = ((i + 1) / totalFiles) * 100;
                    progressBar.style.width = `${progress}%`;
                    progressBar.setAttribute('aria-valuenow', Math.round(progress).toString());
                    console.log(`Progress: ${Math.round(progress)}%`);

                    await new Promise(resolve => requestAnimationFrame(resolve));
                }

                console.log('Finished processing all files. Resetting UI.');
                this.resetProcessingUI();

                if (processedPoemCount > 0) {
                    this.updateDisplay();
                    let message = `Successfully processed ${processedPoemCount} new poem${processedPoemCount > 1 ? 's' : ''}!`;
                    if (skippedCount > 0) {
                        message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                    }
                    this.showNotification(message, 'success');
                    this.announceToScreenReader('process-status', `${processedPoemCount} poems processed successfully`);
                    this.resetFileInput();
                    console.log('Poem processing complete. Display updated.');
                } else {
                    let message = 'No new poems found in the uploaded documents!';
                    if (skippedCount > 0) {
                        message = `All uploaded poems were duplicates or had no new content.`;
                    }
                    this.showNotification(message, 'warning');
                    this.announceToScreenReader('process-status', 'No new poems found');
                    console.log('No new poems added after processing.');
                }

                if (errors.length > 0) {
                    console.error('Summary of processing errors:', errors);
                    this.showNotification(`${errors.length} file(s) had errors. Check console for details.`, 'error', 8000);
                }

            } catch (error) {
                this.resetProcessingUI();
                console.error('Unhandled critical error during document processing:', error);
                this.showNotification('A critical error occurred: ' + error.message, 'error');
                this.announceToScreenReader('process-status', 'Critical error during document processing.');
            }
        }

        /**
         * Resets the UI elements related to document processing.
         */
        resetProcessingUI() {
            console.log('Resetting processing UI.');
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');

            progressContainer.style.display = 'none';
            processBtn.textContent = 'Process Documents';
            processBtn.disabled = this.selectedFiles.length === 0;
            this.isProcessing = false;
        }

        /**
         * Clears the selected files from the input and resets the file label.
         */
        resetFileInput() {
            console.log('Resetting file input.');
            const wordFiles = document.getElementById('wordFiles');
            if (wordFiles) {
                wordFiles.value = '';
                this.handleFileSelect({ target: { files: [] } });
            }
        }

        /**
         * Clears all loaded poems and updates the display.
         */
        clearAllPoems() {
            console.log('Clearing all poems.');
            this.poems = [];
            this.updateDisplay();
            this.resetFileInput();
            this.showNotification('All poems cleared!', 'info');
            this.announceToScreenReader('process-status', 'All poems cleared.');
        }

        /**
         * Extracts HTML content from a DOCX file using Mammoth.js
         * and attempts to identify multiple poems within it.
         * @param {File} file - The DOCX file to process.
         * @returns {Promise<Array<Object>>} A promise resolving to an array of poem objects.
         * @throws {Error} If Mammoth.js is not loaded or content extraction fails.
         */
        async extractPoemsFromDocument(file) {
            console.log(`Attempting to extract poems from "${file.name}"...`);
            if (!window.mammoth) {
                console.error('Mammoth library (window.mammoth) is not loaded.');
                throw new Error('Mammoth library not loaded. Please check the script tag.');
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                console.log(`File "${file.name}" converted to ArrayBuffer.`);
                const result = await window.mammoth.convertToHtml({ arrayBuffer });
                console.log(`Mammoth conversion result for "${file.name}":`, result);

                if (!result.value) {
                    console.warn(`Mammoth returned no HTML content for "${file.name}".`);
                    throw new Error('No content extracted from document by Mammoth.');
                }

                const html = result.value;
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = html;
                const fullContent = tempDiv.textContent.trim();
                console.log(`Full plain text content length for "${file.name}": ${fullContent.length}`);

                if (!fullContent || fullContent.length < 10) {
                    console.warn(`Document "${file.name}" appears empty or too short after extraction.`);
                    throw new Error('Document appears to be empty or too short after extraction.');
                }

                const poems = this.identifyMultiplePoems(tempDiv, file.name, html);
                console.log(`identifyMultiplePoems returned ${poems.length} poems for "${file.name}".`);

                if (poems.length === 0) {
                    const singlePoem = this.createSinglePoemFromDocument(tempDiv, file.name, html, fullContent);
                    console.log(`No multiple poems detected, treating "${file.name}" as a single poem: "${singlePoem.title}".`);
                    return [singlePoem];
                }

                return poems;

            } catch (error) {
                console.error(`Failed to extract content from "${file.name}":`, error);
                throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
            }
        }

        /**
         * Attempts to identify and separate multiple poems within an HTML document structure.
         * Uses different strategies (headings, paragraph breaks, separators).
         * @param {HTMLElement} tempDiv - A temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @param {string} fullHtml - The full HTML content from Mammoth.js.
         * @returns {Array<Object>} An array of identified poem objects.
         */
        identifyMultiplePoems(tempDiv, filename, fullHtml) {
            console.log(`Starting identifyMultiplePoems for "${filename}".`);
            // Strategy 1: Split by headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            if (headings.length > 1) {
                const extractedPoems = this.extractPoemsByHeadings(tempDiv, filename, headings);
                if (extractedPoems.length > 1) {
                    console.log(`Strategy 1 (Headings) found ${extractedPoems.length} poems.`);
                    return extractedPoems;
                }
            }

            // Strategy 2: Split by multiple line breaks or page breaks (empty paragraphs)
            const paragraphs = Array.from(tempDiv.querySelectorAll('p'));
            if (paragraphs.length > 3) {
                const extractedPoems = this.extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs);
                if (extractedPoems.length > 1) {
                    console.log(`Strategy 2 (Paragraph Separation) found ${extractedPoems.length} poems.`);
                    return extractedPoems;
                }
            }

            // Strategy 3: Split by patterns like "***", "---", or similar visual separators
            const textContent = tempDiv.textContent;
            const separatorPatterns = [
                /\n\s*\*{3,}\s*\n/g,
                /\n\s*-{3,}\s*\n/g,
                /\n\s*_{3,}\s*\n/g,
                /\n\s*={3,}\s*\n/g,
                /\n\s*~{3,}\s*\n/g,
                /\n\s*\n\s*\n\s*\n/g
            ];

            for (const pattern of separatorPatterns) {
                const partsHtml = fullHtml.split(pattern);
                if (partsHtml.length > 1) {
                    const extractedPoems = this.extractPoemsBySeparator(partsHtml, filename);
                    if (extractedPoems.length > 1) {
                        console.log(`Strategy 3 (Separators: ${pattern}) found ${extractedPoems.length} poems.`);
                        return extractedPoems;
                    }
                }
            }

            console.log(`No multiple poem strategies yielded results for "${filename}".`);
            return [];
        }

        /**
         * Extracts poems by identifying text blocks separated by heading tags (h1, h2, h3).
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} headings - A NodeList of h1, h2, h3 elements.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsByHeadings(tempDiv, filename, headings) {
            const poems = [];
            const allElements = Array.from(tempDiv.children);
            console.log(`  Extracting by headings for "${filename}". Found ${headings.length} headings.`);

            for (let i = 0; i < headings.length; i++) {
                const currentHeading = headings[i];
                const nextHeading = headings[i + 1];

                const title = currentHeading.textContent.trim() || `Poem ${i + 1} from ${filename}`;

                const startIndex = allElements.indexOf(currentHeading);
                const endIndex = nextHeading ? allElements.indexOf(nextHeading) : allElements.length;

                const poemElements = allElements.slice(startIndex + 1, endIndex);
                const poemContent = poemElements.map(el => el.textContent).join('\n').trim();
                const poemHtml = poemElements.map(el => el.outerHTML).join('\n');

                if (poemContent.length > 10) {
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                } else {
                    console.log(`    Skipping heading "${title}" due to insufficient content.`);
                }
            }
            console.log(`  Finished heading extraction. Found ${poems.length} poems.`);
            return poems.length > 1 ? poems : [];
        }

        /**
         * Extracts poems by identifying blocks of paragraphs separated by empty or very short paragraphs.
         * Attempts to infer titles from the first paragraph of a new block if it fits title criteria.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} paragraphs - A NodeList of paragraph elements.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs) {
            const poems = [];
            let currentPoemElements = [];
            let currentTitle = '';
            let poemIndex = 1;
            console.log(`  Extracting by paragraph separation for "${filename}". Found ${paragraphs.length} paragraphs.`);

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const text = p.textContent.trim();

                const mightBeTitle = text.length > 0 && text.length < 100 &&
                    (p.querySelector('strong') || p.querySelector('b') ||
                     p.style.textAlign === 'center' || /^[A-Z][^.!?]*$/.test(text));

                const isEmptyOrBreak = text.length === 0 || (text.length < 10 && currentPoemElements.length > 0);

                if (isEmptyOrBreak) {
                    if (currentPoemElements.length > 0) {
                        const poemContent = currentPoemElements.map(el => el.textContent).join('\n').trim();
                        const poemHtml = currentPoemElements.map(el => el.outerHTML).join('\n');
                        const title = currentTitle || `Poem ${poemIndex} from ${filename}`;

                        if (poemContent.length > 10) {
                            poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                            poemIndex++;
                            console.log(`    Poem #${poemIndex - 1} identified by paragraph break: "${title}"`);
                        } else {
                            console.log(`    Skipping short poem segment before break.`);
                        }
                        currentPoemElements = [];
                        currentTitle = '';
                    }
                } else if (mightBeTitle && currentPoemElements.length === 0) {
                    currentTitle = text;
                    currentPoemElements.push(p);
                    console.log(`    Potential title detected: "${text}"`);
                } else {
                    currentPoemElements.push(p);
                }
            }

            if (currentPoemElements.length > 0) {
                const poemContent = currentPoemElements.map(el => el.textContent).join('\n').trim();
                const poemHtml = currentPoemElements.map(el => el.outerHTML).join('\n');
                const title = currentTitle || `Poem ${poemIndex} from ${filename}`;

                if (poemContent.length > 10) {
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                    console.log(`    Last poem identified: "${title}"`);
                } else {
                    console.log(`    Skipping last poem segment due to insufficient content.`);
                }
            }
            console.log(`  Finished paragraph separation. Found ${poems.length} poems.`);
            return poems.length > 1 ? poems : [];
        }

        /**
         * Extracts poems by splitting the HTML content based on detected separator patterns.
         * @param {Array<string>} htmlParts - Array of HTML strings separated by a pattern.
         * @param {string} filename - The name of the original file.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsBySeparator(htmlParts, filename) {
            const poems = [];
            console.log(`  Extracting by custom separators for "${filename}". Found ${htmlParts.length} parts.`);
            htmlParts.forEach((part, index) => {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = part.trim();
                const content = tempDiv.textContent.trim();

                if (content.length > 10) {
                    const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                    const firstLine = lines[0] || '';
                    const title = (firstLine.length > 0 && firstLine.length < 100) ?
                        firstLine : `Poem ${index + 1} from ${filename}`;

                    poems.push(this.createPoemObject(title, content, part.trim(), filename));
                    console.log(`    Poem #${index + 1} identified by separator: "${title}"`);
                } else {
                    console.log(`    Skipping part ${index + 1} due to insufficient content after separator.`);
                }
            });
            console.log(`  Finished separator extraction. Found ${poems.length} poems.`);
            return poems;
        }

        /**
         * Creates a single poem object from an entire document when multiple poems are not detected.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The original filename.
         * @param {string} html - The full HTML content from Mammoth.js.
         * @param {string} content - The full plain text content of the document.
         * @returns {Object} A single poem object.
         */
        createSinglePoemFromDocument(tempDiv, filename, html, content) {
            console.log(`Creating single poem object for "${filename}".`);
            const title = this.extractTitle(tempDiv, filename);
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

            return {
                id: Date.now() + Math.random(),
                title: title,
                content: content,
                htmlContent: html,
                filename: filename,
                wordCount: wordCount,
                dateAdded: new Date().toISOString()
            };
        }

        /**
         * Creates a poem object with all necessary properties.
         * @param {string} title - The title of the poem.
         * @param {string} content - The plain text content of the poem.
         * @param {string} htmlContent - The HTML content of the poem.
         * @param {string} filename - The original filename from which the poem was extracted.
         * @returns {Object} The poem object.
         */
        createPoemObject(title, content, htmlContent, filename) {
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;
            return {
                id: Date.now() + Math.random(),
                title: title,
                content: content,
                htmlContent: htmlContent,
                filename: filename,
                wordCount: wordCount,
                dateAdded: new Date().toISOString()
            };
        }

        /**
         * Extracts a title from the document HTML, using various heuristics.
         * @param {HTMLElement} tempDiv - The temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @returns {string} The extracted or generated title.
         */
        extractTitle(tempDiv, filename) {
            let title = '';
            console.log(`Attempting to extract title for "${filename}".`);

            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                if (hText.length > 0 && hText.length < 150) {
                    title = hText;
                    console.log(`  Title found from heading: "${title}"`);
                    break;
                }
            }

            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                for (let i = 0; i < Math.min(3, paragraphs.length); i++) {
                    const p = paragraphs[i];
                    const pText = p.textContent.trim();
                    if (pText.length > 0 && pText.length < 150) {
                        const isBold = p.querySelector('strong, b') !== null;
                        const isCentered = p.style.textAlign === 'center';

                        if (isBold || isCentered) {
                            title = pText;
                            console.log(`  Title found from bold/centered paragraph: "${title}"`);
                            break;
                        }
                    }
                }
            }

            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                if (paragraphs.length > 0) {
                    const firstParagraphText = paragraphs[0].textContent.trim();
                    if (firstParagraphText.length > 0) {
                        const firstLine = firstParagraphText.split('\n')[0].trim();
                        if (firstLine.length > 0 && firstLine.length < 150) {
                            title = firstLine;
                            console.log(`  Title found from first line of first paragraph: "${title}"`);
                        }
                    }
                }
            }

            if (!title) {
                title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
                console.log(`  Title falling back to cleaned filename: "${title}"`);
            }

            title = title.replace(/\s+/g, ' ').trim();
            if (title.length > 150) {
                title = title.substring(0, 147) + '...';
            }

            if (!title) {
                title = "Untitled Poem";
                console.log(`  Title defaulted to "Untitled Poem".`);
            }
            return title;
        }

        /**
         * Updates the display of loaded poems and their count.
         * Attaches drag-and-drop and button event listeners to each poem element.
         */
        updateDisplay() {
            console.log('Updating display. Current poem count:', this.poems.length);
            const poemList = document.getElementById('poemList');
            const poemCountSpan = document.getElementById('poemCount');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');

            if (!poemList || !poemCountSpan || !downloadBtn || !clearBtn) {
                console.error('Required DOM elements for display not found. Please ensure all IDs are correct in the HTML.');
                return;
            }

            poemCountSpan.textContent = this.poems.length;

            if (this.poems.length === 0) {
                poemList.innerHTML = '<div class="no-poems">No poems loaded yet. Upload Word documents to get started!</div>';
                downloadBtn.disabled = true;
                clearBtn.disabled = true;
                return;
            }

            downloadBtn.disabled = false;
            clearBtn.disabled = false;

            poemList.innerHTML = this.poems.map((poem, index) => `
                <div class="poem-item" data-index="${index}" draggable="true">
                    <div class="poem-header">
                        <h3 class="poem-title">${this.escapeHtml(poem.title)}</h3>
                        <div class="poem-meta">
                            <span class="word-count">${poem.wordCount} words</span>
                            <span class="source-file">${this.escapeHtml(poem.filename)}</span>
                        </div>
                        <button class="remove-btn" data-index="${index}" title="Remove this poem">×</button>
                    </div>
                    <div class="poem-content">${poem.content.substring(0, 200)}${poem.content.length > 200 ? '...' : ''}</div>
                </div>
            `).join('');

            // Add drag and drop event listeners
            this.addDragAndDropListeners();
            
            // Add remove button event listeners
            this.addRemoveButtonListeners();
        }

        /**
         * Adds drag and drop event listeners to poem items for reordering.
         */
        addDragAndDropListeners() {
            const poemItems = document.querySelectorAll('.poem-item');
            
            poemItems.forEach((item, index) => {
                item.addEventListener('dragstart', (e) => {
                    this.draggedIndex = index;
                    e.dataTransfer.effectAllowed = 'move';
                    item.classList.add('dragging');
                });

                item.addEventListener('dragend', () => {
                    item.classList.remove('dragging');
                    this.draggedIndex = null;
                });

                item.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                });

                item.addEventListener('drop', (e) => {
                    e.preventDefault();
                    if (this.draggedIndex !== null && this.draggedIndex !== index) {
                        this.reorderPoems(this.draggedIndex, index);
                    }
                });
            });
        }

        /**
         * Adds event listeners to remove buttons for each poem.
         */
        addRemoveButtonListeners() {
            const removeButtons = document.querySelectorAll('.remove-btn');
            removeButtons.forEach(button => {
                button.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const index = parseInt(button.dataset.index);
                    this.removePoem(index);
                });
            });
        }

        /**
         * Reorders poems by moving a poem from one index to another.
         * @param {number} fromIndex - The original index of the poem.
         * @param {number} toIndex - The target index for the poem.
         */
        reorderPoems(fromIndex, toIndex) {
            const poem = this.poems.splice(fromIndex, 1)[0];
            this.poems.splice(toIndex, 0, poem);
            this.updateDisplay();
            this.showNotification('Poems reordered!', 'info');
        }

        /**
         * Removes a poem at the specified index.
         * @param {number} index - The index of the poem to remove.
         */
        removePoem(index) {
            if (index >= 0 && index < this.poems.length) {
                const removedPoem = this.poems.splice(index, 1)[0];
                this.updateDisplay();
                this.showNotification(`Removed "${removedPoem.title}"`, 'info');
                this.announceToScreenReader('process-status', `Poem "${removedPoem.title}" removed`);
            }
        }

        /**
         * Downloads all poems as a combined Word document.
         */
        async downloadCombinedDocument() {
            if (this.poems.length === 0) {
                this.showNotification('No poems to download!', 'warning');
                return;
            }

            try {
                let combinedHtml = `
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>Combined Poems</title>
                        <style>
                            body { font-family: 'Times New Roman', serif; line-height: 1.6; margin: 40px; }
                            .poem { margin-bottom: 50px; page-break-after: auto; }
                            .poem-title { font-size: 18px; font-weight: bold; text-align: center; margin-bottom: 20px; }
                            .poem-content { white-space: pre-wrap; }
                            .page-break { page-break-before: always; }
                        </style>
                    </head>
                    <body>
                `;

                this.poems.forEach((poem, index) => {
                    if (index > 0) {
                        combinedHtml += '<div class="page-break"></div>';
                    }
                    combinedHtml += `
                        <div class="poem">
                            <div class="poem-title">${this.escapeHtml(poem.title)}</div>
                            <div class="poem-content">${poem.htmlContent || this.escapeHtml(poem.content)}</div>
                        </div>
                    `;
                });

                combinedHtml += '</body></html>';

                const blob = new Blob([combinedHtml], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `combined-poems-${new Date().toISOString().split('T')[0]}.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);

                this.showNotification(`Downloaded ${this.poems.length} poems successfully!`, 'success');
                this.announceToScreenReader('process-status', `${this.poems.length} poems downloaded`);

            } catch (error) {
                console.error('Error creating download:', error);
                this.showNotification('Error creating download: ' + error.message, 'error');
            }
        }

        /**
         * Escapes HTML characters to prevent XSS attacks.
         * @param {string} text - The text to escape.
         * @returns {string} The escaped text.
         */
        escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        /**
         * Shows a notification message to the user.
         * @param {string} message - The message to display.
         * @param {string} type - The type of notification (success, error, warning, info).
         * @param {number} duration - How long to show the notification in milliseconds (default: 5000).
         */
        showNotification(message, type = 'info', duration = 5000) {
            if (this.notificationTimeout) {
                clearTimeout(this.notificationTimeout);
            }

            const notification = document.getElementById('notification');
            if (!notification) {
                console.warn('Notification element not found.');
                return;
            }

            notification.textContent = message;
            notification.className = `notification ${type} show`;

            if (duration > 0) {
                this.notificationTimeout = setTimeout(() => {
                    notification.classList.remove('show');
                }, duration);
            }
        }

        /**
         * Announces messages to screen readers for accessibility.
         * @param {string} ariaLiveId - The ID of the aria-live element.
         * @param {string} message - The message to announce.
         */
        announceToScreenReader(ariaLiveId, message) {
            const ariaLive = document.getElementById(ariaLiveId);
            if (ariaLive) {
                ariaLive.textContent = message;
                setTimeout(() => {
                    ariaLive.textContent = '';
                }, 3000);
            }
        }
    }

    // Initialize the application when the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
        console.log('DOM loaded, initializing PoemCompiler...');
        new PoemCompiler();
    });

})();
