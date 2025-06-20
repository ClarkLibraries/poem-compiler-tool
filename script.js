// Ensure strict mode for better code quality
'use strict';

(function() {
    class PoemCompiler {
        constructor() {
            this.poems = []; // Array to store processed poems
            this.notificationTimeout = null; // To manage notification timeouts

            // Announce initial state for screen readers (optional)
            this.announceToScreenReader('download-status', 'Poem Compiler tool loaded.');

            // Make the compiler instance globally accessible if needed for HTML events
            // In a real app, you'd bind events more robustly, e.g., with addEventListener
            window.poemCompiler = this;
        }

        // Placeholder for your actual poem processing logic
        processDocument(poemContent) {
            // In a real scenario, you'd parse, compile, and format the poem
            // For now, let's just add it to the poems array
            const newPoem = `<div><h2>Poem ${this.poems.length + 1}</h2><pre>${poemContent}</pre></div>`;
            this.poems.push(newPoem);
            this.showNotification('Document processed and poem added!', 'success');
            this.announceToScreenReader('download-status', 'Poem processed.');
        }

        // Generates the combined HTML for download
        generateCombinedHTML() {
            if (this.poems.length === 0) {
                return '<html><body><h1>No Poems Processed</h1><p>Process some poems first!</p></body></html>';
            }
            let html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Combined Poems Collection</title>
    <style>
        body { font-family: sans-serif; margin: 40px; line-height: 1.6; }
        h1 { color: #333; }
        h2 { color: #555; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-top: 30px; }
        pre { background-color: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; }
    </style>
</head>
<body>
    <h1>Your Combined Poem Collection</h1>`;
            this.poems.forEach((poem, index) => {
                html += `<div>${poem}</div>`;
            });
            html += `</body>
</html>`;
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
