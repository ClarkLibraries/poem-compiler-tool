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
            if (paragraphs.length > 3) { // Ensure there are enough paragraphs to consider separation
                const extractedPoems = this.extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs);
                if (extractedPoems.length > 1) {
                    console.log(`Strategy 2 (Paragraph Separation) found ${extractedPoems.length} poems.`);
                    return extractedPoems;
                }
            }

            // Strategy 3: Split by patterns like "***", "---", or similar visual separators
            const separatorPatterns = [
                /\n\s*\*{3,}\s*\n/g, // ***
                /\n\s*-{3,}\s*\n/g, // ---
                /\n\s*_{3,}\s*\n/g, // ___
                /\n\s*={3,}\s*\n/g, // ===
                /\n\s*~{3,}\s*\n/g, // ~~~
                /(<p>\s*&nbsp;\s*<\/p>){2,}/g, // Multiple empty paragraphs (mammoth often produces &nbsp;)
                /(<p>\s*<\/p>){2,}/g // Multiple empty paragraphs
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

                // Collect all elements between current heading and the next (or end of document)
                const poemElements = allElements.slice(startIndex + 1, endIndex);
                
                // Filter out any empty text nodes or very short paragraphs that might be artifacts
                const meaningfulElements = poemElements.filter(el => el.textContent.trim().length > 0 || el.tagName === 'BR');

                if (meaningfulElements.length === 0) {
                    console.log(`    Skipping heading "${title}" as no content found before next heading/end.`);
                    continue;
                }

                const poemContent = meaningfulElements.map(el => {
                    if (el.tagName === 'BR') return '\n'; // Preserve line breaks
                    return el.textContent;
                }).join('\n').trim();

                const poemHtml = meaningfulElements.map(el => el.outerHTML).join('\n');

                if (poemContent.length > 10) { // Ensure there's substantial content
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
                const html = p.outerHTML;

                // Heuristic for what might be a title: short, possibly bold/centered, starts with a capital letter
                const mightBeTitle = text.length > 0 && text.length < 100 &&
                    (p.querySelector('strong') || p.querySelector('b') ||
                    p.style.textAlign === 'center' || /^[A-Z]/.test(text)); // Starts with a capital letter

                // Heuristic for an empty line or a significant break: empty paragraph or very short and not clearly content
                const isEmptyOrBreak = text.length === 0 || (text.length < 5 && currentPoemElements.length > 0 && !mightBeTitle);

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
                    // If multiple empty paragraphs, consider it a strong separator, do not add the empty p to content
                } else if (mightBeTitle && currentPoemElements.length === 0) {
                    currentTitle = text;
                    // Add this potential title paragraph to the current poem elements, as it's part of its structure
                    currentPoemElements.push(p);
                    console.log(`    Potential title detected: "${text}"`);
                } else {
                    currentPoemElements.push(p);
                }
            }

            // Add any remaining poem elements after the loop finishes
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

                // If the part is just the separator itself or very short, skip it
                if (content.length < 10) {
                    console.log(`    Skipping part ${index + 1} due to insufficient content after separator.`);
                    return;
                }

                // Attempt to find a title within this part, prioritizing headings or bold/centered text
                let title = '';
                const headings = tempDiv.querySelectorAll('h1, h2, h3');
                if (headings.length > 0) {
                    title = headings[0].textContent.trim();
                } else {
                    const paragraphs = tempDiv.querySelectorAll('p');
                    if (paragraphs.length > 0) {
                        const firstParaText = paragraphs[0].textContent.trim();
                        if (firstParaText.length > 0 && firstParaText.length < 150) {
                            const isBold = paragraphs[0].querySelector('strong, b') !== null;
                            const isCentered = paragraphs[0].style.textAlign === 'center';
                            if (isBold || isCentered) {
                                title = firstParaText;
                            }
                        }
                    }
                }

                if (!title) {
                    // Fallback to first non-empty line as title if no clear title found
                    const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                    const firstLine = lines[0] || '';
                    if (firstLine.length > 0 && firstLine.length < 100) {
                        title = firstLine;
                    } else {
                        title = `Poem ${index + 1} from ${filename}`;
                    }
                }

                poems.push(this.createPoemObject(title, content, part.trim(), filename));
                console.log(`    Poem #${index + 1} identified by separator: "${title}"`);
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

            // Prioritize actual heading tags
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                if (hText.length > 0 && hText.length < 150) {
                    title = hText;
                    console.log(`  Title found from heading: "${title}"`);
                    break;
                }
            }

            // If no heading, check first few paragraphs for bold/centered text
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

            // As a last resort, use the first non-empty line of content or filename
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

            return title;
        }

        /**
         * Updates the display of loaded poems in the UI.
         * Enables/disables the download button based on poem count.
         * Re-initializes drag-and-drop for poem reordering.
         */
        updateDisplay() {
            console.log('Updating display for poems. Total poems:', this.poems.length);
            const poemsList = document.getElementById('poemsList');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');
            const placeholder = document.getElementById('poemsPlaceholder');
            const poemCountSpan = document.getElementById('poemCount');

            if (!poemsList || !downloadBtn || !clearBtn || !placeholder || !poemCountSpan) {
                console.error('Required DOM elements for display update not found.');
                this.showNotification('Application display error: Some UI elements are missing. Please check the HTML.', 'error', 0);
                return;
            }

            poemsList.innerHTML = ''; // Clear existing list
            poemCountSpan.textContent = this.poems.length.toString();

            if (this.poems.length === 0) {
                placeholder.style.display = 'block';
                poemsList.style.display = 'none';
                downloadBtn.disabled = true;
                clearBtn.disabled = true;
                this.announceToScreenReader('poem-list-status', 'No poems loaded.');
                console.log('No poems to display. Placeholder shown, buttons disabled.');
                return;
            }

            placeholder.style.display = 'none';
            poemsList.style.display = 'block';
            downloadBtn.disabled = false;
            clearBtn.disabled = false;

            this.poems.forEach((poem, index) => {
                const li = document.createElement('li');
                li.className = 'poem-item bg-white p-4 shadow-sm rounded-lg flex items-center justify-between transition-all duration-200 ease-in-out';
                li.draggable = true;
                li.dataset.id = poem.id;
                li.dataset.index = index;

                li.innerHTML = `
                    <div class="flex-1 min-w-0">
                        <h3 class="text-lg font-semibold text-gray-800 truncate">${this.escapeHtml(poem.title)}</h3>
                        <p class="text-sm text-gray-500 truncate">${this.escapeHtml(poem.filename)} - ${poem.wordCount} words</p>
                    </div>
                    <div class="flex items-center space-x-2 ml-4">
                        <button type="button" class="view-poem-btn p-2 rounded-full text-blue-600 hover:bg-blue-100 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50" aria-label="View poem ${this.escapeHtml(poem.title)}" data-id="${poem.id}">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                <path d="M10 12a2 2 0 100-4 2 2 0 000 4z" />
                                <path fill-rule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z" clip-rule="evenodd" />
                            </svg>
                        </button>
                        <button type="button" class="remove-poem-btn p-2 rounded-full text-red-600 hover:bg-red-100 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-opacity-50" aria-label="Remove poem ${this.escapeHtml(poem.title)}" data-id="${poem.id}">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                <path fill-rule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm6 0a1 1 0 012 0v6a1 1 0 11-2 0V8z" clip-rule="evenodd" />
                            </svg>
                        </button>
                    </div>
                `;
                poemsList.appendChild(li);
            });

            this.addPoemListEventListeners();
            this.announceToScreenReader('poem-list-status', `${this.poems.length} poems loaded. Use drag and drop to reorder.`);
            console.log('Poem list rendered and event listeners added.');
        }

        /**
         * Adds event listeners for viewing, removing, and reordering poems.
         */
        addPoemListEventListeners() {
            const poemsList = document.getElementById('poemsList');
            if (!poemsList) {
                console.error('Poem list element not found for event listeners.');
                return;
            }

            poemsList.querySelectorAll('.view-poem-btn').forEach(button => {
                button.onclick = (e) => this.viewPoem(e.currentTarget.dataset.id);
            });

            poemsList.querySelectorAll('.remove-poem-btn').forEach(button => {
                button.onclick = (e) => this.removePoem(e.currentTarget.dataset.id);
            });

            // Drag and Drop for reordering
            poemsList.addEventListener('dragstart', (e) => {
                const target = e.target.closest('.poem-item');
                if (target) {
                    this.draggedIndex = parseInt(target.dataset.index, 10);
                    e.dataTransfer.effectAllowed = 'move';
                    // Using setTimeout to prevent Firefox drag ghosting issue on the original element
                    setTimeout(() => target.classList.add('dragging'), 0);
                    console.log('Drag started for index:', this.draggedIndex);
                }
            });

            poemsList.addEventListener('dragover', (e) => {
                this.preventDefaults(e); // Allow drop
                const target = e.target.closest('.poem-item');
                if (target && target.dataset.index !== undefined && this.draggedIndex !== null) {
                    const dragOverIndex = parseInt(target.dataset.index, 10);
                    if (this.draggedIndex !== dragOverIndex) {
                        const draggedEl = poemsList.querySelector(`[data-index="${this.draggedIndex}"]`);
                        if (draggedEl) {
                            // Visually reorder elements for feedback
                            if (dragOverIndex > this.draggedIndex) {
                                target.parentNode.insertBefore(draggedEl, target.nextSibling);
                            } else {
                                target.parentNode.insertBefore(draggedEl, target);
                            }
                            // Temporarily update data-index to reflect new visual order for next dragover
                            // This is a common but tricky part of drag-and-drop reordering.
                            // A more robust solution might involve direct manipulation of the underlying array and re-rendering.
                            // For simplicity, we'll rely on the 'drop' event to finalize.
                        }
                    }
                }
            });

            poemsList.addEventListener('dragleave', (e) => {
                // Not strictly necessary for simple reordering, but good practice
                // to remove any visual indicators if dragged item leaves the list area.
            });

            poemsList.addEventListener('drop', (e) => {
                this.preventDefaults(e);
                const target = e.target.closest('.poem-item');
                if (target && this.draggedIndex !== null) {
                    const dropIndex = parseInt(target.dataset.index, 10);
                    console.log(`Drop detected. Dragged Index: ${this.draggedIndex}, Dropped onto Index: ${dropIndex}`);

                    // Finalize the reordering in the actual poems array
                    const [draggedPoem] = this.poems.splice(this.draggedIndex, 1);
                    this.poems.splice(dropIndex, 0, draggedPoem);

                    this.showNotification(`Reordered poem "${draggedPoem.title}"`, 'info');
                    this.announceToScreenReader('poem-list-status', `Poem ${draggedPoem.title} moved to position ${dropIndex + 1}.`);

                    this.draggedIndex = null;
                    // Re-render the list to reflect the new order and update data-index attributes correctly
                    this.updateDisplay();
                }
            });

            poemsList.addEventListener('dragend', (e) => {
                const draggedEl = poemsList.querySelector('.poem-item.dragging');
                if (draggedEl) {
                    draggedEl.classList.remove('dragging');
                }
                this.draggedIndex = null; // Reset
                console.log('Drag ended. draggedIndex reset.');
            });
            console.log('Drag and drop listeners added to poem list.');
        }


        /**
         * Removes a poem from the list by its ID.
         * @param {string} id - The ID of the poem to remove.
         */
        removePoem(id) {
            console.log('Attempting to remove poem with ID:', id);
            const initialCount = this.poems.length;
            this.poems = this.poems.filter(poem => poem.id != id); // Use != for loose comparison with dataset.id (string)
            if (this.poems.length < initialCount) {
                this.showNotification('Poem removed!', 'success');
                this.announceToScreenReader('poem-list-status', 'Poem removed.');
                this.updateDisplay();
                console.log('Poem removed successfully. Remaining poems:', this.poems.length);
            } else {
                console.warn('Poem with ID not found:', id);
            }
        }

        /**
         * Displays a poem's content in a modal.
         * @param {string} id - The ID of the poem to view.
         */
        viewPoem(id) {
            console.log('Viewing poem with ID:', id);
            const poem = this.poems.find(p => p.id == id);
            if (poem) {
                const modal = document.getElementById('poemModal');
                const modalTitle = document.getElementById('poemModalTitle');
                const modalContent = document.getElementById('poemModalContent');
                const closeModalBtn = document.getElementById('closeModal');

                if (!modal || !modalTitle || !modalContent || !closeModalBtn) {
                    console.error('Modal elements not found.');
                    this.showNotification('Error: Modal display elements missing.', 'error');
                    return;
                }

                modalTitle.textContent = poem.title;
                // Use innerHTML to preserve formatting from Mammoth.js
                modalContent.innerHTML = poem.htmlContent;
                modal.classList.remove('hidden');
                modal.setAttribute('aria-hidden', 'false');
                modal.focus(); // Focus the modal for accessibility

                closeModalBtn.onclick = () => {
                    modal.classList.add('hidden');
                    modal.setAttribute('aria-hidden', 'true');
                    // Return focus to the button that opened the modal if possible
                    document.querySelector(`button[data-id="${id}"]`)?.focus();
                };

                // Close modal on escape key
                const handleEscape = (e) => {
                    if (e.key === 'Escape') {
                        closeModalBtn.click();
                        document.removeEventListener('keydown', handleEscape);
                    }
                };
                document.addEventListener('keydown', handleEscape);

                console.log(`Modal opened for "${poem.title}".`);
            } else {
                this.showNotification('Poem not found.', 'error');
                console.warn('Attempted to view non-existent poem ID:', id);
            }
        }

        /**
         * Combines all loaded poems into a single HTML document and triggers a download.
         */
        downloadCombinedDocument() {
            if (this.poems.length === 0) {
                this.showNotification('No poems to download!', 'warning');
                console.warn('Download attempted with no poems.');
                return;
            }

            console.log('Preparing combined document for download.');
            let combinedHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Combined Poems</title>
    <style>
        body { font-family: sans-serif; line-height: 1.6; max-width: 800px; margin: 2em auto; padding: 0 1em; color: #333; }
        h1, h2, h3 { color: #2c3e50; margin-top: 1.5em; margin-bottom: 0.5em; }
        h1 { font-size: 2.2em; text-align: center; border-bottom: 2px solid #eee; padding-bottom: 0.5em; }
        h2 { font-size: 1.8em; }
        h3 { font-size: 1.4em; }
        .poem-section { margin-bottom: 2em; padding-bottom: 1em; border-bottom: 1px dashed #eee; }
        .poem-section:last-child { border-bottom: none; margin-bottom: 0; padding-bottom: 0; }
        p { margin-bottom: 0.5em; }
        .poem-metadata { font-size: 0.9em; color: #777; margin-top: -0.5em; margin-bottom: 1em; }
        pre { background-color: #f4f4f4; padding: 1em; border-radius: 5px; overflow-x: auto; white-space: pre-wrap; word-wrap: break-word; }
        /* Preserve line breaks for poetry if not already in paragraphs */
        .poem-content { white-space: pre-wrap; } 
    </style>
</head>
<body>
    <h1>Combined Poems</h1>
    <div class="poems-container">
`;

            this.poems.forEach(poem => {
                combinedHtml += `
        <div class="poem-section">
            <h2>${this.escapeHtml(poem.title)}</h2>
            <p class="poem-metadata"><em>From: ${this.escapeHtml(poem.filename)} | Words: ${poem.wordCount}</em></p>
            <div class="poem-content">
                ${poem.htmlContent}
            </div>
        </div>
`;
            });

            combinedHtml += `
    </div>
</body>
</html>`;

            const blob = new Blob([combinedHtml], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Combined_Poems.html';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            this.showNotification('Combined document downloaded!', 'success');
            this.announceToScreenReader('process-status', 'Combined document downloaded.');
            console.log('Combined document download initiated.');
        }

        /**
         * Displays a temporary notification message to the user.
         * @param {string} message - The message to display.
         * @param {string} type - The type of notification (e.g., 'success', 'error', 'info', 'warning').
         * @param {number} [duration=5000] - How long the notification should be visible in milliseconds.
         */
        showNotification(message, type, duration = 5000) {
            const notification = document.getElementById('notification');
            if (!notification) {
                console.error('Notification element not found.');
                return;
            }

            // Clear any existing timeout to prevent rapid notifications from hiding too quickly
            if (this.notificationTimeout) {
                clearTimeout(this.notificationTimeout);
                this.notificationTimeout = null;
            }

            notification.textContent = message;
            notification.className = `notification fixed bottom-4 right-4 p-3 rounded-md shadow-lg text-white opacity-0 transition-opacity duration-300 z-50`;

            switch (type) {
                case 'success':
                    notification.classList.add('bg-green-500');
                    break;
                case 'error':
                    notification.classList.add('bg-red-500');
                    break;
                case 'info':
                    notification.classList.add('bg-blue-500');
                    break;
                case 'warning':
                    notification.classList.add('bg-yellow-500');
                    notification.classList.add('text-gray-900'); // Ensure text is visible on yellow
                    break;
                default:
                    notification.classList.add('bg-gray-700');
            }

            // Show notification
            requestAnimationFrame(() => {
                notification.classList.remove('opacity-0');
                notification.classList.add('opacity-100');
            });

            // Hide after duration
            this.notificationTimeout = setTimeout(() => {
                notification.classList.remove('opacity-100');
                notification.classList.add('opacity-0');
                this.notificationTimeout = null;
            }, duration);

            console.log(`Notification: ${message} (${type})`);
        }

        /**
         * Announces messages to screen readers using an ARIA live region.
         * @param {string} regionId - The ID of the live region element.
         * @param {string} message - The message to announce.
         */
        announceToScreenReader(regionId, message) {
            const liveRegion = document.getElementById(regionId);
            if (liveRegion) {
                liveRegion.textContent = message;
                console.log(`Announced to screen reader (${regionId}): ${message}`);
            } else {
                console.warn(`ARIA live region with ID "${regionId}" not found.`);
            }
        }

        /**
         * Escapes HTML entities in a string to prevent XSS.
         * @param {string} str - The string to escape.
         * @returns {string} The escaped string.
         */
        escapeHtml(str) {
            const div = document.createElement('div');
            div.appendChild(document.createTextNode(str));
            return div.innerHTML;
        }
    }

    // Initialize the PoemCompiler once the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
        console.log('DOM Content Loaded. Initializing PoemCompiler.');
        new PoemCompiler();
    });
})();
