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
                let processedPoemCount = 0;
                let skippedCount = 0;
                const totalFiles = this.selectedFiles.length;
                const errors = [];

                for (let i = 0; i < this.selectedFiles.length; i++) {
                    const file = this.selectedFiles[i];

                    try {
                        const poemsFromFile = await this.extractPoemsFromDocument(file);
                        if (poemsFromFile && poemsFromFile.length > 0) {
                            // Check for duplicates and add new poems
                            for (const poemData of poemsFromFile) {
                                if (poemData && poemData.content && poemData.content.trim().length > 0) {
                                    // Check for duplicates based on title and content similarity
                                    const isDuplicate = this.poems.some(existing =>
                                        existing.title.toLowerCase() === poemData.title.toLowerCase() ||
                                        (existing.content.trim().length > 50 && existing.content.trim() === poemData.content.trim())
                                    );

                                    if (!isDuplicate) {
                                        this.poems.push(poemData);
                                        processedPoemCount++;
                                    } else {
                                        skippedCount++;
                                        console.warn(`Duplicate poem detected and skipped: ${poemData.title || 'Untitled'}`);
                                    }
                                }
                            }
                        } else {
                            errors.push(`${file.name}: No valid poems found`);
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
                if (processedPoemCount > 0) {
                    this.updateDisplay();
                    let message = `Successfully processed ${processedPoemCount} new poem${processedPoemCount > 1 ? 's' : ''}!`;
                    if (skippedCount > 0) {
                        message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                    }
                    this.showNotification(message, 'success');
                    this.announceToScreenReader('process-status', `${processedPoemCount} poems processed successfully`);

                    // Reset file input only if some poems were successfully processed
                    this.resetFileInput();
                } else {
                    let message = 'No new poems found in the uploaded documents!';
                    if (skippedCount > 0) {
                        message = `All uploaded poems were duplicates or had no new content.`;
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

        async extractPoemsFromDocument(file) {
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

                // Get full content for analysis
                const fullContent = tempDiv.textContent.trim();

                if (!fullContent || fullContent.length < 10) {
                    throw new Error('Document appears to be empty or too short after extraction.');
                }

                // Extract multiple poems from the document
                const poems = this.identifyMultiplePoems(tempDiv, file.name, html);
                
                if (poems.length === 0) {
                    // If no poems were identified, treat the entire document as one poem
                    const singlePoem = this.createSinglePoemFromDocument(tempDiv, file.name, html, fullContent);
                    return [singlePoem];
                }

                return poems;

            } catch (error) {
                throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
            }
        }

        identifyMultiplePoems(tempDiv, filename, html) {
            const poems = [];
            
            // Strategy 1: Split by headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            if (headings.length > 1) {
                return this.extractPoemsByHeadings(tempDiv, filename, headings);
            }

            // Strategy 2: Split by multiple line breaks or page breaks
            const paragraphs = Array.from(tempDiv.querySelectorAll('p'));
            if (paragraphs.length > 3) {
                return this.extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs, html);
            }

            // Strategy 3: Split by patterns like "***", "---", or similar separators
            const textContent = tempDiv.textContent;
            const separatorPatterns = [
                /\*{3,}/g,           // Three or more asterisks
                /-{3,}/g,            // Three or more dashes
                /_{3,}/g,            // Three or more underscores
                /={3,}/g,            // Three or more equals
                /~{3,}/g,            // Three or more tildes
                /\n\s*\n\s*\n/g      // Three or more line breaks
            ];

            for (const pattern of separatorPatterns) {
                const parts = textContent.split(pattern);
                if (parts.length > 1) {
                    return this.extractPoemsBySeparator(parts, filename, tempDiv, html);
                }
            }

            return []; // Return empty array if no multiple poems detected
        }

        extractPoemsByHeadings(tempDiv, filename, headings) {
            const poems = [];
            const allElements = Array.from(tempDiv.children);
            
            for (let i = 0; i < headings.length; i++) {
                const currentHeading = headings[i];
                const nextHeading = headings[i + 1];
                
                const title = currentHeading.textContent.trim() || `Poem ${i + 1}`;
                
                // Find content between this heading and the next
                const startIndex = allElements.indexOf(currentHeading);
                const endIndex = nextHeading ? allElements.indexOf(nextHeading) : allElements.length;
                
                const poemElements = allElements.slice(startIndex + 1, endIndex);
                const poemContent = poemElements.map(el => el.textContent).join('\n').trim();
                const poemHtml = poemElements.map(el => el.outerHTML).join('\n');
                
                if (poemContent.length > 10) {
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                }
            }
            
            return poems;
        }

        extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs, html) {
            const poems = [];
            let currentPoem = [];
            let currentTitle = '';
            let poemIndex = 1;
            
            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const text = p.textContent.trim();
                
                // Check if this might be a title (short line, possibly centered or bold)
                const mightBeTitle = text.length < 100 && text.length > 0 && 
                    (p.querySelector('strong') || p.querySelector('b') || 
                     p.style.textAlign === 'center' || text.match(/^[A-Z][^.!?]*$/));
                
                // Check for poem separator (empty paragraph or very short paragraph)
                const isEmpty = text.length === 0;
                const isVeryShort = text.length < 5;
                
                if (isEmpty || (isVeryShort && currentPoem.length > 0)) {
                    // End current poem if we have content
                    if (currentPoem.length > 0) {
                        const poemContent = currentPoem.map(el => el.textContent).join('\n').trim();
                        const poemHtml = currentPoem.map(el => el.outerHTML).join('\n');
                        const title = currentTitle || `Poem ${poemIndex}`;
                        
                        if (poemContent.length > 10) {
                            poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                            poemIndex++;
                        }
                        
                        currentPoem = [];
                        currentTitle = '';
                    }
                } else if (mightBeTitle && currentPoem.length === 0) {
                    // This might be a title for the next poem
                    currentTitle = text;
                    currentPoem.push(p);
                } else {
                    // Regular content
                    currentPoem.push(p);
                }
            }
            
            // Handle the last poem
            if (currentPoem.length > 0) {
                const poemContent = currentPoem.map(el => el.textContent).join('\n').trim();
                const poemHtml = currentPoem.map(el => el.outerHTML).join('\n');
                const title = currentTitle || `Poem ${poemIndex}`;
                
                if (poemContent.length > 10) {
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                }
            }
            
            return poems.length > 1 ? poems : []; // Only return if we found multiple poems
        }

        extractPoemsBySeparator(parts, filename, tempDiv, html) {
            const poems = [];
            
            parts.forEach((part, index) => {
                const content = part.trim();
                if (content.length > 10) {
                    // Try to extract a title from the first line
                    const lines = content.split('\n');
                    const firstLine = lines[0].trim();
                    const title = (firstLine.length < 100 && firstLine.length > 0) ? 
                        firstLine : `Poem ${index + 1}`;
                    
                    // For HTML content, we'll use a simplified version since we split by text
                    const simpleHtml = content.split('\n').map(line => 
                        `<p>${this.escapeHtml(line)}</p>`).join('\n');
                    
                    poems.push(this.createPoemObject(title, content, simpleHtml, filename));
                }
            });
            
            return poems;
        }

        createSinglePoemFromDocument(tempDiv, filename, html, content) {
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

        extractTitle(tempDiv, filename) {
            let title = '';

            // 1. Try headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                if (hText.length > 0 && hText.length < 150) {
                    title = hText;
                    break;
                }
            }

            // 2. Try bold or centered text at the beginning
            if (!title) {
                const boldElements = tempDiv.querySelectorAll('strong, b');
                for (let i = 0; i < boldElements.length; i++) {
                    const bText = boldElements[i].textContent.trim();
                    if (bText.length > 0 && bText.length < 150) {
                        title = bText;
                        break;
                    }
                }
            }

            // 3. If no clear heading, try first line
            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                if (paragraphs.length > 0) {
                    const firstParagraphText = paragraphs[0].textContent.trim();
                    if (firstParagraphText.length > 0) {
                        const firstLine = firstParagraphText.split('\n')[0].trim();
                        if (firstLine.length > 0 && firstLine.length < 150) {
                            title = firstLine;
                        }
                    }
                }
            }

            // 4. Fallback to filename (cleaned up)
            if (!title) {
                title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
            }

            // Final cleanup
            title = title.replace(/\s+/g, ' ').trim();
            if (title.length > 150) {
                title = title.substring(0, 147) + '...';
            }

            if (!title) {
                title = "Untitled Poem";
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
            poemDiv.setAttribute('tabindex', '0');

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
            if (typeof text !== 'string') return '';
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
        return '<h2>Table of Contents</h2><p><em>No poems available</em></p>';
    }

    let tocHtml = '<h2>Table of Contents</h2>\n<ul class="toc-list">\n';
    
    this.poems.forEach((poem, index) => {
        const pageNumber = index + 1; // Simple page numbering
        const safeTitle = this.escapeHtml(poem.title);
        const safeFilename = this.escapeHtml(poem.filename);
        
        tocHtml += `  <li class="toc-item">
    <span class="toc-title">
      <a href="#poem-${index}" class="toc-link">${safeTitle}</a>
    </span>
    <span class="toc-dots">............................................</span>
    <span class="toc-page">${pageNumber}</span>
    <br>
    <small class="toc-source">Source: ${safeFilename} (${poem.wordCount} words)</small>
  </li>\n`;
    });
    
    tocHtml += '</ul>\n';
    
    // Add some basic CSS for the table of contents
    tocHtml += `
<style>
.toc-list {
    list-style: none;
    padding: 0;
    margin: 20px 0;
}

.toc-item {
    margin-bottom: 12px;
    padding: 8px 0;
    border-bottom: 1px solid #eee;
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
}

.toc-title {
    flex: 0 0 auto;
    font-weight: bold;
}

.toc-link {
    color: #333;
    text-decoration: none;
}

.toc-link:hover {
    color: #0066cc;
    text-decoration: underline;
}

.toc-dots {
    flex: 1 1 auto;
    text-align: center;
    color: #ccc;
    font-family: monospace;
    overflow: hidden;
    white-space: nowrap;
    margin: 0 8px;
}

.toc-page {
    flex: 0 0 auto;
    font-weight: bold;
    min-width: 30px;
    text-align: right;
}

.toc-source {
    flex: 0 0 100%;
    color: #666;
    font-style: italic;
    margin-top: 4px;
}

@media print {
    .toc-dots {
        display: none;
    }
    
    .toc-item {
        page-break-inside: avoid;
    }
}
</style>`;
    
    return tocHtml;
}

downloadCombinedDocument() {
    if (this.poems.length === 0) {
        this.showNotification('No poems to download!', 'warning');
        return;
    }

    try {
        // Generate the complete document
        let documentHtml = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Compiled Poems Collection</title>
    <style>
        body {
            font-family: 'Times New Roman', serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 40px 20px;
            color: #333;
        }
        
        h1 {
            text-align: center;
            border-bottom: 2px solid #333;
            padding-bottom: 20px;
            margin-bottom: 40px;
        }
        
        .poem {
            margin-bottom: 60px;
            page-break-inside: avoid;
            border-bottom: 1px solid #eee;
            padding-bottom: 40px;
        }
        
        .poem-title {
            font-size: 1.5em;
            font-weight: bold;
            margin-bottom: 10px;
            text-align: center;
        }
        
        .poem-meta {
            font-size: 0.9em;
            color: #666;
            text-align: center;
            margin-bottom: 20px;
            font-style: italic;
        }
        
        .poem-content {
            white-space: pre-line;
            margin: 20px 0;
        }
        
        @media print {
            body {
                padding: 20px;
            }
            
            .poem {
                page-break-after: always;
            }
            
            .poem:last-child {
                page-break-after: auto;
            }
        }
    </style>
</head>
<body>
    <h1>Compiled Poems Collection</h1>
    <p style="text-align: center; font-style: italic; margin-bottom: 40px;">
        Generated on ${new Date().toLocaleDateString()} • ${this.poems.length} poem${this.poems.length > 1 ? 's' : ''}
    </p>
`;

        // Add table of contents
        documentHtml += this.generateTableOfContents();
        documentHtml += '\n<div style="page-break-before: always;"></div>\n';

        // Add each poem
        this.poems.forEach((poem, index) => {
            documentHtml += `
    <div class="poem" id="poem-${index}">
        <h2 class="poem-title">${this.escapeHtml(poem.title)}</h2>
        <div class="poem-meta">
            Source: ${this.escapeHtml(poem.filename)} • ${poem.wordCount} words
        </div>
        <div class="poem-content">
            ${poem.htmlContent || this.escapeHtml(poem.content)}
        </div>
    </div>
`;
        });

        documentHtml += `
</body>
</html>`;

        // Create and download the file
        const blob = new Blob([documentHtml], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `compiled-poems-${new Date().toISOString().split('T')[0]}.html`;
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        this.showNotification(`Downloaded ${this.poems.length} poems successfully!`, 'success');
        this.announceToScreenReader('download-status', `Downloaded compilation of ${this.poems.length} poems`);

    } catch (error) {
        console.error('Download error:', error);
        this.showNotification('Error creating download: ' + error.message, 'error');
    }
}

showNotification(message, type = 'info', duration = 4000) {
    // Clear any existing notification timeout
    if (this.notificationTimeout) {
        clearTimeout(this.notificationTimeout);
    }

    const notification = document.getElementById('notification');
    if (!notification) {
        console.warn('Notification element not found');
        return;
    }

    notification.textContent = message;
    notification.className = `widget-notification show ${type}`;

    this.notificationTimeout = setTimeout(() => {
        notification.className = 'widget-notification';
    }, duration);
}

announceToScreenReader(elementId, message) {
    const element = document.getElementById(elementId);
    if (element) {
        element.textContent = message;
    } else {
        // Create a temporary live region if the element doesn't exist
        const liveRegion = document.createElement('div');
        liveRegion.setAttribute('aria-live', 'polite');
        liveRegion.setAttribute('aria-atomic', 'true');
        liveRegion.style.position = 'absolute';
        liveRegion.style.left = '-10000px';
        liveRegion.style.width = '1px';
        liveRegion.style.height = '1px';
        liveRegion.style.overflow = 'hidden';
        liveRegion.textContent = message;
        
        document.body.appendChild(liveRegion);
        setTimeout(() => {
            if (liveRegion.parentNode) {
                liveRegion.parentNode.removeChild(liveRegion);
            }
        }, 1000);
    }
}
