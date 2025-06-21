(function() {
        'use strict';

        class PoemCompiler {
            constructor() {
                this.poems = [];
                this.selectedFiles = [];
                this.draggedIndex = null;
                this.isProcessing = false;
                this.notificationTimeout = null;
                this.initializeEventListeners();
                this.updateDisplay();
            }

            initializeEventListeners() {
                const wordFiles = document.getElementById('wordFiles');
                const processBtn = document.getElementById('processBtn');
                const downloadBtn = document.getElementById('downloadBtn');
                const clearBtn = document.getElementById('clearBtn');
                const fileLabel = document.getElementById('fileLabel');

                if (!wordFiles || !processBtn || !downloadBtn || !clearBtn || !fileLabel) {
                    console.error('Required DOM elements not found');
                    return;
                }

                wordFiles.addEventListener('change', (e) => {
                    this.handleFileSelect(e);
                });

                processBtn.addEventListener('click', () => {
                    if (!this.isProcessing) {
                        this.processDocuments();
                    }
                });

                downloadBtn.addEventListener('click', () => {
                    this.downloadCombinedDocument();
                });

                clearBtn.addEventListener('click', () => {
                    this.clearAllPoems();
                });

                // Enhanced drag and drop functionality for the file label
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
                    const files = Array.from(e.dataTransfer.files).filter(file =>
                        file.name.toLowerCase().endsWith('.docx')
                    );
                    if (files.length > 0) {
                        this.setFileInput(files);
                        // Trigger change event programmatically to update the file input
                        const event = new Event('change', { bubbles: true });
                        wordFiles.dispatchEvent(event);
                    } else if (e.dataTransfer.files.length > 0) {
                        this.showNotification('Please upload only .docx files', 'warning');
                    }
                }, false);
            }

            preventDefaults(e) {
                e.preventDefault();
                e.stopPropagation();
            }

            setFileInput(files) {
                try {
                    const dt = new DataTransfer();
                    files.forEach(file => dt.items.add(file));
                    document.getElementById('wordFiles').files = dt.files;
                } catch (error) {
                    console.warn('Could not set file input directly (might be IE/Edge):', error);
                    // Fallback: just store the files reference if DataTransfer fails
                    this.selectedFiles = files;
                }
            }

            handleFileSelect(event) {
                const files = Array.from(event.target.files);
                const fileLabel = document.getElementById('fileLabel');
                const processBtn = document.getElementById('processBtn');

                // Validate file types
                const validFiles = files.filter(file => file.name.toLowerCase().endsWith('.docx'));
                const invalidFiles = files.filter(file => !file.name.toLowerCase().endsWith('.docx'));

                if (invalidFiles.length > 0) {
                    this.showNotification(`${invalidFiles.length} invalid file(s) ignored. Only .docx files are supported.`, 'warning');
                }

                if (validFiles.length > 0) {
                    this.selectedFiles = validFiles;
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
                }
            }

            async processDocuments() {
                if (this.selectedFiles.length === 0) {
                    this.showNotification('Please select Word documents first!', 'warning');
                    return;
                }

                if (this.isProcessing) {
                    return; // Prevent multiple simultaneous processing
                }

                this.isProcessing = true;
                const processBtn = document.getElementById('processBtn');
                const progressContainer = document.getElementById('progressContainer');
                const progressBar = document.getElementById('progressBar');

                // Update UI for processing state
                processBtn.disabled = true;
                processBtn.textContent = 'Processing...';
                progressContainer.style.display = 'block';
                progressBar.style.width = '0%';

                this.announceToScreenReader('process-status', 'Processing documents...');

                try {
                    let processedCount = 0;
                    let skippedCount = 0;
                    const totalFiles = this.selectedFiles.length;
                    const errors = [];

                    for (let i = 0; i < this.selectedFiles.length; i++) {
                        const file = this.selectedFiles[i];

                        try {
                            const poemData = await this.extractPoemFromDocument(file);
                            if (poemData && poemData.content && poemData.content.trim().length > 0) {
                                // Check for duplicates based on title and content similarity
                                const isDuplicate = this.poems.some(existing =>
                                    existing.title.toLowerCase() === poemData.title.toLowerCase() ||
                                    (existing.content.trim().length > 50 && existing.content.trim() === poemData.content.trim())
                                );

                                if (!isDuplicate) {
                                    this.poems.push(poemData);
                                    processedCount++;
                                } else {
                                    skippedCount++;
                                    console.warn(`Duplicate poem detected and skipped: ${poemData.title || file.name}`);
                                    this.showNotification(`Skipped duplicate: "${poemData.title || file.name}"`, 'info');
                                }
                            } else {
                                errors.push(`${file.name}: No valid content found`);
                            }
                        } catch (error) {
                            console.error(`Error processing ${file.name}:`, error);
                            errors.push(`${file.name}: ${error.message}`);
                        }

                        // Update progress
                        const progress = ((i + 1) / totalFiles) * 100;
                        progressBar.style.width = `${progress}%`;

                        // Small delay to show progress, but not block UI
                        await new Promise(resolve => requestAnimationFrame(resolve));
                    }

                    // Reset UI state
                    this.resetProcessingUI();

                    // Show results
                    if (processedCount > 0) {
                        this.updateDisplay();
                        let message = `Successfully processed ${processedCount} new poem${processedCount > 1 ? 's' : ''}!`;
                        if (skippedCount > 0) {
                            message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                        }
                        this.showNotification(message, 'success');
                        this.announceToScreenReader('process-status', `${processedCount} poems processed successfully`);

                        // Reset file input only if some files were successfully processed
                        this.resetFileInput();
                    } else {
                        let message = 'No new poems found in the uploaded documents!';
                        if (skippedCount > 0) {
                            message = `All ${skippedCount} uploaded documents were duplicates or had no new content.`;
                        }
                        this.showNotification(message, 'warning');
                        this.announceToScreenReader('process-status', 'No new poems found');
                    }

                    // Show errors if any
                    if (errors.length > 0) {
                        console.error('Processing errors:', errors);
                        this.showNotification(`${errors.length} file(s) had errors. Check console for details.`, 'error', 8000);
                    }

                } catch (error) {
                    this.resetProcessingUI();
                    console.error('Document processing error:', error);
                    this.showNotification('Error processing documents: ' + error.message, 'error');
                    this.announceToScreenReader('process-status', 'Error processing documents');
                }
            }

            resetProcessingUI() {
                const processBtn = document.getElementById('processBtn');
                const progressContainer = document.getElementById('progressContainer');

                progressContainer.style.display = 'none';
                processBtn.textContent = 'Process Documents';
                processBtn.disabled = this.selectedFiles.length === 0;
                this.isProcessing = false;
            }

            resetFileInput() {
                const wordFiles = document.getElementById('wordFiles');
                if (wordFiles) {
                    wordFiles.value = ''; // Clears the selected file(s) from the input
                    this.handleFileSelect({ target: { files: [] } }); // Resets the label text and button state
                }
            }

            async extractPoemFromDocument(file) {
                if (!window.mammoth) {
                    throw new Error('Mammoth library not loaded. Please check the script tag.');
                }

                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const result = await window.mammoth.convertToHtml({ arrayBuffer });

                    if (!result.value) {
                        throw new Error('No content extracted from document by Mammoth.');
                    }

                    const html = result.value;
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = html;

                    // Extract title - improved logic
                    let title = this.extractTitle(tempDiv, file.name);

                    // Get full content, strip HTML tags for plain text content
                    const content = tempDiv.textContent.trim();

                    if (!content || content.length < 10) {
                        throw new Error('Document appears to be empty or too short after extraction.');
                    }

                    // Calculate word count more accurately
                    const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

                    return {
                        id: Date.now() + Math.random(), // Unique ID for each poem
                        title: title,
                        content: content,
                        htmlContent: html, // Keep HTML content for combined document
                        filename: file.name,
                        wordCount: wordCount,
                        dateAdded: new Date().toISOString()
                    };
                } catch (error) {
                    throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
                }
            }

            extractTitle(tempDiv, filename) {
                let title = '';

                // 1. Try headings (H1, H2)
                const headings = tempDiv.querySelectorAll('h1, h2');
                for (let i = 0; i < headings.length; i++) {
                    const hText = headings[i].textContent.trim();
                    if (hText.length > 0 && hText.length < 150) { // Limit length for reasonable titles
                        title = hText;
                        break;
                    }
                }

                // 2. If no clear heading, try first few lines of text
                if (!title) {
                    const paragraphs = tempDiv.querySelectorAll('p');
                    if (paragraphs.length > 0) {
                        const firstParagraphText = paragraphs[0].textContent.trim();
                        if (firstParagraphText.length > 0) {
                            // Take the first line as a potential title if it's reasonably short
                            const firstLine = firstParagraphText.split('\n')[0].trim();
                            if (firstLine.length > 0 && firstLine.length < 150) {
                                title = firstLine;
                            }
                        }
                    }
                }

                // 3. Fallback to filename (cleaned up)
                if (!title) {
                    title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
                }

                // Final cleanup
                title = title.replace(/\s+/g, ' ').trim(); // Replace multiple spaces with single space
                if (title.length > 150) {
                    title = title.substring(0, 147) + '...'; // Truncate if too long
                }

                if (!title) {
                    title = "Untitled Poem"; // Default if all else fails
                }

                return title;
            }

            updateDisplay() {
                const poemList = document.getElementById('poemList');
                const poemCountSpan = document.getElementById('poemCount');
                const downloadBtn = document.getElementById('downloadBtn');
                const clearBtn = document.getElementById('clearBtn');

                if (!poemList || !poemCountSpan || !downloadBtn || !clearBtn) {
                    console.error('Required display elements not found for updateDisplay');
                    return;
                }

                poemList.innerHTML = '';
                poemCountSpan.textContent = this.poems.length;

                if (this.poems.length === 0) {
                    poemList.innerHTML = `
                        <div class="widget-empty-state">
                            <p>No poems loaded yet. Upload and process Word documents to begin.</p>
                        </div>
                    `;
                    downloadBtn.disabled = true;
                    clearBtn.disabled = true;
                } else {
                    this.poems.forEach((poem, index) => {
                        const poemDiv = this.createPoemElement(poem, index);
                        poemList.appendChild(poemDiv);
                    });

                    downloadBtn.disabled = false;
                    clearBtn.disabled = false;
                }
            }

            createPoemElement(poem, index) {
                const poemDiv = document.createElement('div');
                poemDiv.classList.add('widget-poem-item');
                poemDiv.setAttribute('draggable', 'true');
                poemDiv.setAttribute('data-index', index);
                poemDiv.setAttribute('role', 'listitem');
                poemDiv.setAttribute('aria-label', `Poem: ${poem.title}, position ${index + 1} of ${this.poems.length}`);
                poemDiv.setAttribute('tabindex', '0'); // Make draggable items focusable for accessibility

                // Create safe preview text
                const preview = poem.content.length > 100
                    ? poem.content.substring(0, 100) + '...'
                    : poem.content;

                poemDiv.innerHTML = `
                    <div class="widget-drag-indicator" aria-hidden="true">⋮⋮</div>
                    <div class="widget-poem-details">
                        <h3>${this.escapeHtml(poem.title)}</h3>
                        <p><strong>Source:</strong> ${this.escapeHtml(poem.filename)}</p>
                        <p><strong>Word Count:</strong> ${poem.wordCount}</p>
                        <p><strong>Preview:</strong> ${this.escapeHtml(preview)}</p>
                    </div>
                    <div class="widget-poem-controls">
                        ${index > 0 ? `<button class="widget-move-btn"
                                data-index="${index}"
                                aria-label="Move ${this.escapeHtml(poem.title)} up in the list">
                            <span aria-hidden="true">↑</span>
                        </button>` : '<div style="width: 32px; visibility: hidden;"></div>'}
                        <button class="widget-remove-btn"
                                data-index="${index}"
                                aria-label="Remove ${this.escapeHtml(poem.title)} from the list">
                            <span aria-hidden="true">×</span>
                        </button>
                        ${index < this.poems.length - 1 ? `<button class="widget-move-btn move-down"
                                data-index="${index}"
                                aria-label="Move ${this.escapeHtml(poem.title)} down in the list">
                            <span aria-hidden="true">↓</span>
                        </button>` : '<div style="width: 32px; visibility: hidden;"></div>'}
                    </div>
                `;

                this.attachPoemEventListeners(poemDiv, index);
                return poemDiv;
            }

            attachPoemEventListeners(poemDiv, index) {
                // Drag and drop event listeners
                poemDiv.addEventListener('dragstart', (e) => {
                    this.draggedIndex = index;
                    poemDiv.classList.add('dragging');
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', index.toString());
                    this.announceToScreenReader('process-status', `Started dragging ${this.poems[index].title}`);
                });

                poemDiv.addEventListener('dragend', () => {
                    poemDiv.classList.remove('dragging');
                    this.draggedIndex = null;
                });

                poemDiv.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                    const targetElement = e.currentTarget;
                    if (targetElement.classList.contains('widget-poem-item') && this.draggedIndex !== null) {
                        const targetIndex = parseInt(targetElement.dataset.index);
                        if (targetIndex !== this.draggedIndex) {
                            poemDiv.classList.add('drag-over');
                        }
                    }
                });

                poemDiv.addEventListener('dragleave', (e) => {
                    e.currentTarget.classList.remove('drag-over');
                });

                poemDiv.addEventListener('drop', (e) => {
                    e.preventDefault();
                    e.currentTarget.classList.remove('drag-over');
                    const draggedIdx = parseInt(e.dataTransfer.getData('text/plain'));
                    const dropTargetIndex = parseInt(e.currentTarget.dataset.index);

                    if (draggedIdx !== dropTargetIndex && !isNaN(draggedIdx)) {
                        this.movePoem(draggedIdx, dropTargetIndex);
                    }
                });

                // Keyboard accessibility for drag and drop
                poemDiv.addEventListener('keydown', (e) => {
                    if (e.key === 'ArrowUp' && e.ctrlKey && index > 0) {
                        e.preventDefault();
                        this.movePoem(index, index - 1);
                        // Re-focus the moved item to maintain accessibility
                        requestAnimationFrame(() => {
                            const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index - 1}"]`);
                            if (newPoemDiv) newPoemDiv.focus();
                        });
                    } else if (e.key === 'ArrowDown' && e.ctrlKey && index < this.poems.length - 1) {
                        e.preventDefault();
                        this.movePoem(index, index + 1);
                        // Re-focus the moved item
                        requestAnimationFrame(() => {
                            const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index + 1}"]`);
                            if (newPoemDiv) newPoemDiv.focus();
                        });
                    } else if (e.key === 'Delete' || e.key === 'Backspace') {
                        e.preventDefault();
                        this.removePoem(index);
                    }
                });


                // Move up button
                const moveUpBtn = poemDiv.querySelector('.widget-move-btn:not(.move-down)');
                if (moveUpBtn) {
                    moveUpBtn.addEventListener('click', (e) => {
                        e.preventDefault();
                        if (index > 0) {
                            this.movePoem(index, index - 1);
                        }
                    });
                }

                // Move down button
                const moveDownBtn = poemDiv.querySelector('.widget-move-btn.move-down');
                if (moveDownBtn) {
                    moveDownBtn.addEventListener('click', (e) => {
                        e.preventDefault();
                        if (index < this.poems.length - 1) {
                            this.movePoem(index, index + 1);
                        }
                    });
                }


                // Remove button
                const removeBtn = poemDiv.querySelector('.widget-remove-btn');
                if (removeBtn) {
                    removeBtn.addEventListener('click', (e) => {
                        e.preventDefault();
                        this.removePoem(index);
                    });
                }
            }

            escapeHtml(text) {
                if (typeof text !== 'string') return ''; // Ensure text is a string
                const div = document.createElement('div');
                div.textContent = text;
                return div.innerHTML;
            }

            movePoem(fromIndex, toIndex) {
                if (fromIndex < 0 || fromIndex >= this.poems.length ||
                    toIndex < 0 || toIndex >= this.poems.length ||
                    fromIndex === toIndex) {
                    return;
                }

                const poem = this.poems.splice(fromIndex, 1)[0];
                this.poems.splice(toIndex, 0, poem);
                this.updateDisplay();
                this.announceToScreenReader('process-status', `Moved "${poem.title}" to position ${toIndex + 1} of ${this.poems.length}`);
                this.showNotification(`Moved "${poem.title}"`, 'info');
            }

            removePoem(index) {
                if (index < 0 || index >= this.poems.length) {
                    return;
                }

                const poem = this.poems[index];
                const confirmed = confirm(`Are you sure you want to remove "${poem.title}"?`);

                if (confirmed) {
                    this.poems.splice(index, 1);
                    this.updateDisplay();
                    this.showNotification(`Removed "${poem.title}"`, 'info');
                    this.announceToScreenReader('process-status', `Removed poem: ${poem.title}. There are now ${this.poems.length} poems.`);
                }
            }

            clearAllPoems() {
                if (this.poems.length === 0) return;

                const confirmed = confirm(`Are you sure you want to clear all ${this.poems.length} poems? This action cannot be undone.`);
                if (confirmed) {
                    this.poems = [];
                    this.updateDisplay();
                    this.showNotification('All poems cleared', 'info');
                    this.announceToScreenReader('clear-status', 'All poems cleared.');
                }
            }

            generateTableOfContents() {
                if (this.poems.length === 0) {
                    return '';
                }
                let toc = '<h1>Table of Contents</h1>\n<nav role="navigation" aria-label="Table of Contents">\n<ul>\n';
                this.poems.forEach((poem, index) => {
                    const id = `poem-${index + 1}`;
                    toc += `<li><a href="#${id}">${this.escapeHtml(poem.title)}</a></li>\n`;
                });
                toc += '</ul>\n</nav>\n<hr>\n\n';
                return toc;
            }

            generateCombinedHTML() {
                const totalWords = this.poems.reduce((sum, poem) => sum + poem.wordCount, 0);
                const generatedDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

                let html = '<!DOCTYPE html>\n';
                html += '<html lang="en">\n';
                html += '<head>\n';
                html += '<meta charset="UTF-8">\n';
                html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n';
                html += '<title>Combined Poems Collection</title>\n';
                html += '<style>\n';
                html += 'body {\n';
                html += '    font-family: \'Times New Roman\', serif;\n';
                html += '    line-height: 1.6;\n';
                html += '    max-width: 800px;\n';
                html += '    margin: 0 auto;\n';
                html += '    padding: 40px 20px;\n';
                html += '    color: #333;\n';
                html += '    background-color: #fff;\n';
                html += '}\n';
                html += 'h1 {\n';
                html += '    color: #2c3e50;\n';
                html += '    text-align: center;\n';
                html += '    margin-bottom: 30px;\n';
                html += '    font-size: 2.5em;\n';
                html += '    border-bottom: 3px solid #3498db;\n';
                html += '    padding-bottom: 15px;\n';
                html += '}\n';
                html += 'h2 {\n';
                html += '    color: #34495e;\n';
                html += '    margin-top: 60px;\n';
                html += '    margin-bottom: 20px;\n';
                html += '    font-size: 1.8em;\n';
                html += '    border-bottom: 2px solid #ecf0f1;\n';
                html += '    padding-bottom: 10px;\n';
                html += '}\n';
                html += '.collection-info {\n';
                html += '    text-align: center;\n';
                html += '    margin-bottom: 40px;\n';
                html += '    padding: 20px;\n';
                html += '    background: #f8f9fa;\n';
                html += '    border-radius: 8px;\n';
                html += '    border-left: 4px solid #3498db;\n';
                html += '}\n';
                html += '.collection-info p {\n';
                html += '    margin: 5px 0;\n';
                html += '    color: #666;\n';
                html += '}\n';
                html += 'nav ul {\n'; // Added nav for TOC
                html += '    list-style-type: none;\n';
                html += '    padding: 0;\n';
                html += '}\n';
                html += 'nav li {\n'; // Added nav for TOC
                html += '    margin: 10px 0;\n';
                html += '    padding: 8px 0;\n';
                html += '    border-bottom: 1px dotted #ddd;\n';
                html += '}\n';
                html += 'a {\n';
                html += '    color: #3498db;\n';
                html += '    text-decoration: none;\n';
                html += '    font-weight: 500;\n';
                html += '    padding: 5px;\n';
                html += '    border-radius: 3px;\n';
                html += '    transition: background-color 0.2s;\n';
                html += '}\n';
                html += 'a:hover {\n';
                html += '    background-color: #e3f2fd;\n';
                html += '    text-decoration: underline;\n';
                html += '}\n';
                html += 'hr {\n';
                html += '    border: none;\n';
                html += '    height: 2px;\n';
                html += '    background: linear-gradient(to right, #3498db, #ecf0f1, #3498db);\n';
                html += '    margin: 40px 0;\n';
                html += '}\n';
                html += '.poem-content {\n';
                html += '    margin-bottom: 50px;\n';
                html += '    padding: 25px;\n';
                html += '    background: #fafafa;\n';
                html += '    border-left: 4px solid #3498db;\n';
                html += '    border-radius: 0 8px 8px 0;\n';
                html += '    box-shadow: 0 2px 4px rgba(0,0,0,0.1);\n';
                html += '}\n';
                html += '.poem-meta {\n';
                html += '    font-size: 0.9em;\n';
                html += '    color: #7f8c8d;\n';
                html += '    margin-bottom: 20px;\n';
                html += '    font-style: italic;\n';
                html += '    padding: 10px;\n';
                html += '    background: #ecf0f1;\n';
                html += '    border-radius: 4px;\n';
                html += '}\n';
                html += '.poem-text {\n';
                html += '    font-size: 1.1em;\n';
                html += '    line-height: 1.8;\n';
                html += '}\n';
                html += '@media print {\n';
                html += '    body { \n';
                html += '        padding: 20px;\n';
                html += '        font-size: 12pt;\n';
                html += '    }\n';
                html += '    .poem-content { \n';
                html += '        page-break-inside: avoid;\n';
                html += '        box-shadow: none;\n';
                html += '    }\n';
                html += '    .collection-info {\n';
                html += '        background: transparent;\n';
                html += '    }\n';
                html += '}\n';
                html += '@media (max-width: 600px) {\n';
                html += '    body {\n';
                html += '        padding: 20px 15px;\n';
                html += '    }\n';
                html += '    h1 {\n';
                html += '        font-size: 2em;\n';
                html += '    }\n';
                html += '    .poem-content {\n';
                html += '        padding: 15px;\n';
                html += '    }\n';
                html += '}\n';
                html += '</style>\n';
                html += '</head>\n';
                html += '<body>\n';
                html += '<h1>Combined Poems Collection</h1>\n';
                html += '<div class="collection-info">\n';
                html += '<p><strong>Total Poems:</strong> ' + this.poems.length + '</p>\n';
                html += '<p><strong>Total Words:</strong> ' + totalWords.toLocaleString() + '</p>\n';
                html += '<p><strong>Generated:</strong> ' + generatedDate + '</p>\n';
                html += '</div>\n';

                // Add table of contents
                html += this.generateTableOfContents();

                // Add each poem
                this.poems.forEach((poem, index) => {
                    html += '<div class="poem-content">\n';
                    html += '<h2 id="poem-' + (index + 1) + '">' + this.escapeHtml(poem.title) + '</h2>\n';
                    html += '<div class="poem-meta">\n';
                    html += '<strong>Source:</strong> ' + this.escapeHtml(poem.filename) + ' | \n';
                    html += '<strong>Word Count:</strong> ' + poem.wordCount.toLocaleString() + ' | \n';
                    html += '<strong>Added:</strong> ' + new Date(poem.dateAdded).toLocaleDateString() + '\n';
                    html += '</div>\n';
                    html += '<div class="poem-text">\n';
                    html += poem.htmlContent + '\n'; // Use htmlContent directly for richer formatting
                    html += '</div>\n';
                    html += '</div>\n';
                });

                html += '</body>\n';
                html += '</html>';
                return html;
            }


            downloadCombinedDocument() {
                if (this.poems.length === 0) {
                    this.showNotification('No poems to download. Please process documents first.', 'warning');
                    return;
                }
                this.announceToScreenReader('download-status', 'Starting download of combined document.');

                try {
                    const combinedHtml = this.generateCombinedHTML();
                    const blob = new Blob([combinedHtml], { type: 'text/html' });
                    const url = URL.createObjectURL(blob);

                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'Combined_Poems_Collection.html';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);

                    this.showNotification('Combined document downloaded successfully!', 'success');
                    this.announceToScreenReader('download-status', 'Combined document download complete.');
                } catch (error) {
                    console.error('Error downloading combined document:', error);
                    this.showNotification('Error downloading document: ' + error.message, 'error');
                    this.announceToScreenReader('download-status', 'Error during document download.');
                }
            }

            // --- Notification and Accessibility Utilities ---

            showNotification(message, type = 'info', duration = 5000) {
                let notificationContainer = document.getElementById('notificationContainer');
                if (!notificationContainer) {
                    notificationContainer = document.createElement('div');
                    notificationContainer.id = 'notificationContainer';
                    Object.assign(notificationContainer.style, {
                        position: 'fixed',
                        bottom: '20px',
                        left: '50%',
                        transform: 'translateX(-50%)',
                        zIndex: '1000',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '10px',
                        maxWidth: '90%',
                        pointerEvents: 'none' // Allows clicks to pass through
                    });
                    document.body.appendChild(notificationContainer);
                }

                const notification = document.createElement('div');
                notification.classList.add('widget-notification', `widget-notification-${type}`);
                notification.textContent = message;
                notification.setAttribute('role', 'alert');
                notification.setAttribute('aria-live', 'polite');

                Object.assign(notification.style, {
                    padding: '12px 20px',
                    borderRadius: '8px',
                    color: '#fff',
                    textAlign: 'center',
                    boxShadow: '0 4px 12px rgba(0,0,0,0.2)',
                    opacity: '0',
                    transition: 'opacity 0.3s ease-in-out, transform 0.3s ease-in-out',
                    transform: 'translateY(20px)',
                    pointerEvents: 'auto'
                });

                if (type === 'success') {
                    notification.style.backgroundColor = '#28a745';
                } else if (type === 'error') {
                    notification.style.backgroundColor = '#dc3545';
                } else if (type === 'warning') {
                    notification.style.backgroundColor = '#ffc107';
                    notification.style.color = '#333';
                } else { // info
                    notification.style.backgroundColor = '#17a2b8';
                }

                notificationContainer.appendChild(notification);

                // Animate in
                requestAnimationFrame(() => {
                    notification.style.opacity = '1';
                    notification.style.transform = 'translateY(0)';
                });

                // Clear any existing timeout for new notification
                if (this.notificationTimeout) {
                    clearTimeout(this.notificationTimeout);
                }

                // Animate out and remove after duration
                this.notificationTimeout = setTimeout(() => {
                    notification.style.opacity = '0';
                    notification.style.transform = 'translateY(20px)';
                    notification.addEventListener('transitionend', () => {
                        notification.remove();
                        if (notificationContainer.children.length === 0) {
                            notificationContainer.remove();
                        }
                    }, { once: true });
                }, duration);
            }

            announceToScreenReader(elementId, message) {
                const statusElement = document.getElementById(elementId);
                if (statusElement) {
                    // Clear existing content to ensure re-announcement if message is the same
                    statusElement.textContent = '';
                    // Set timeout to ensure screen reader detects change
                    setTimeout(() => {
                        statusElement.textContent = message;
                    }, 100);
                }
            }
        }

        // Initialize the compiler when the DOM is fully loaded
        document.addEventListener('DOMContentLoaded', () => {
            new PoemCompiler();
        });
    })();
