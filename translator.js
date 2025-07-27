document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const translateBtn = document.getElementById('translateBtn');
    const clearCacheBtn = document.getElementById('clearCacheBtn');
    const progressSection = document.querySelector('.progress-section');
    const progressBar = document.getElementById('progressBar');
    const status = document.getElementById('status');
    const apiCalls = document.getElementById('apiCalls');
    const cached = document.getElementById('cached');
    const logContainer = document.getElementById('logContainer');
    const summarySection = document.querySelector('.summary-section');
    const summaryDisplay = document.getElementById('summaryDisplay');
    const summaryEditor = document.getElementById('summaryEditor');
    const editSummaryBtn = document.getElementById('editSummaryBtn');
    const saveSummaryBtn = document.getElementById('saveSummaryBtn');
    const exportSummaryBtn = document.getElementById('exportSummaryBtn');

    let currentSummary = '';

    // Summary editing handlers
    editSummaryBtn.addEventListener('click', () => {
        summaryEditor.value = summaryDisplay.textContent;
        summaryDisplay.style.display = 'none';
        summaryEditor.style.display = 'block';
        editSummaryBtn.style.display = 'none';
        saveSummaryBtn.style.display = 'inline-flex';
    });

    saveSummaryBtn.addEventListener('click', () => {
        currentSummary = summaryEditor.value;
        summaryDisplay.textContent = currentSummary;
        summaryEditor.style.display = 'none';
        summaryDisplay.style.display = 'block';
        saveSummaryBtn.style.display = 'none';
        editSummaryBtn.style.display = 'inline-flex';
    });

    exportSummaryBtn.addEventListener('click', async () => {
        try {
            const response = await fetch('/export-summary', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ summary: currentSummary }),
            });

            if (!response.ok) {
                throw new Error('Failed to export summary');
            }

            // Trigger download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'document_summary.docx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            addLog('Summary exported as DOCX successfully');
        } catch (error) {
            addLog('Error exporting summary: ' + error.message);
        }
    });

    // Drag and drop handlers
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    dropZone.addEventListener('drop', handleDrop, false);

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight(e) {
        dropZone.classList.add('dragover');
    }

    function unhighlight(e) {
        dropZone.classList.remove('dragover');
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }

    // File input change handler
    fileInput.addEventListener('change', function(e) {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (file.name.toLowerCase().endsWith('.docx') || file.name.toLowerCase().endsWith('.pdf')) {
                fileName.textContent = file.name;
                translateBtn.disabled = false;
                addLog('File selected: ' + file.name);
            } else {
                fileName.textContent = 'Please select a DOCX or PDF file';
                translateBtn.disabled = true;
                addLog('Error: Invalid file type. Please select a DOCX or PDF file.');
            }
        }
    }

    // Translation process
    translateBtn.addEventListener('click', async () => {
        const file = fileInput.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);
        formData.append('targetLang', document.getElementById('targetLang').value);
        formData.append('engine', document.getElementById('engine').value);
        formData.append('generateSummary', document.getElementById('generateSummary').checked);
        formData.append('firstPageOnly', document.getElementById('firstPageOnly').checked);
        formData.append('tone', document.getElementById('toneSelect').value);
        formData.append('pdfEngine', document.getElementById('pdfEngine').value);

        progressSection.style.display = 'block';
        translateBtn.disabled = true;
        status.textContent = 'Translating...';
        progressBar.style.width = '0%';
        
        let currentFile = file.name;
        
        // Clear log container and add initial message
        logContainer.innerHTML = '';
        addLog('Starting translation process...');

        try {
            addLog('Sending file to server...');
            const response = await fetch('/translate', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error(`Translation failed with status: ${response.status}`);
            }
            
            addLog('Connected to server stream. Waiting for updates...');

            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            
            let translationComplete = false;
            let summaryGenerated = false;
            let finalDownloadUrl = null;

            while (true) {
                const {value, done} = await reader.read();
                if (done) {
                    addLog('Stream complete');
                    break;
                }

                const text = decoder.decode(value);
                
                const updates = text.split('\n').filter(line => line.trim());

                updates.forEach(update => {
                    try {
                        const data = JSON.parse(update);
                        handleProgressUpdate(data);
                        
                        if (data.downloadUrl) {
                            finalDownloadUrl = data.downloadUrl;
                            translationComplete = true;
                            logger.info('Received final download URL:', finalDownloadUrl);
                        }
                        
                        if (data.summary) {
                            summaryGenerated = true;
                            currentSummary = data.summary;
                            summaryDisplay.textContent = data.summary;
                            summarySection.style.display = 'block';
                            editSummaryBtn.disabled = false;
                            exportSummaryBtn.disabled = false;
                            addLog('Document summary generated and displayed');
                        }
                    } catch (e) {
                        addLog(`Non-JSON update: ${update}`);
                    }
                });
            }

            if (translationComplete && finalDownloadUrl) {
                status.textContent = 'Translation completed';
                progressBar.style.width = '100%';
                addDownloadButton(finalDownloadUrl);
                
                addLog('Translation process finished successfully. Download ready.');
            } else {
                if (status.textContent !== 'Error') {
                    status.textContent = 'Translation incomplete or failed';
                }
                addLog('Translation process finished with errors or was incomplete.');
            }
        } catch (error) {
            status.textContent = 'Error';
            addLog(`Error during translation: ${error.message}`);
            console.error('Translation error:', error);
        } finally {
            translateBtn.disabled = false;
        }
    });

    // Clear cache button handler
    clearCacheBtn.addEventListener('click', async () => {
        try {
            const response = await fetch('/clear-cache', {
                method: 'POST'
            });

            if (!response.ok) {
                throw new Error('Failed to clear cache');
            }

            const data = await response.json();
            apiCalls.textContent = '0';
            cached.textContent = '0';
            addLog('Cache cleared: ' + data.message);
        } catch (error) {
            addLog('Error clearing cache: ' + error.message);
        }
    });

    // Helper functions
    function handleProgressUpdate(data) {
        if (data.progress !== undefined) {
            progressBar.style.width = data.progress + '%';
        }
        if (data.status) {
            status.textContent = data.status;
        }
        if (data.apiCalls !== undefined) {
            apiCalls.textContent = data.apiCalls;
        }
        if (data.cached !== undefined) {
            cached.textContent = data.cached;
        }
        if (data.message) {
            addLog(data.message);
        }
        if (data.downloadUrl) {
            console.log("Download URL received:", data.downloadUrl);
        }
    }

    function addLog(message) {
        const timestamp = new Date().toLocaleTimeString();
        const logEntry = document.createElement('div');
        logEntry.className = 'log-entry';
        logEntry.textContent = `[${timestamp}] ${message}`;
        logContainer.appendChild(logEntry);
        logContainer.scrollTop = logContainer.scrollHeight;
    }

    // Add download button function
    function addDownloadButton(downloadUrl) {
        let downloadBtn = document.getElementById('downloadBtn');
        
        if (!downloadBtn) {
            downloadBtn = document.createElement('button');
            downloadBtn.id = 'downloadBtn';
            downloadBtn.className = 'success-btn';
            downloadBtn.innerHTML = '<i class="fas fa-download"></i> Download Translated Document';
            
            const actionButtons = document.querySelector('.action-buttons');
            actionButtons.prepend(downloadBtn);
        }
        
        downloadBtn.onclick = function() {
            window.location.href = downloadUrl;
            addLog('Downloading translated document...');
        };
        
        addLog('Download ready! Click the download button to save the translated document.');
    }

    // Add summary download button function
    function addSummaryDownloadButton(filename) {
        let summaryBtn = document.getElementById('summaryBtn');
        
        if (!summaryBtn) {
            summaryBtn = document.createElement('button');
            summaryBtn.id = 'summaryBtn';
            summaryBtn.className = 'secondary-btn';
            summaryBtn.innerHTML = '<i class="fas fa-file-alt"></i> Download Summary';
            
            const actionButtons = document.querySelector('.action-buttons');
            actionButtons.appendChild(summaryBtn);
        }
        
        summaryBtn.onclick = function() {
            window.location.href = `/download-summary/${encodeURIComponent(filename)}`;
            addLog('Downloading document summary...');
        };
        
        addLog('Summary ready! Click the summary button to download.');
    }
}); 
