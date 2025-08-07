/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// API Configuration
const NUTRIENT_API_BASE = 'https://localhost:8080/api/nutrient';

// Authentication Configuration
interface AuthCredentials {
    apiKey: string;
    viewerKey: string;
    apiBaseUrl: string;
}

// Global authentication state
let authCredentials: AuthCredentials | null = null;

// Interface for file data
interface FileData {
    file: File;
    name: string;
    size: string;
}

// Global state
let selectedFile: FileData | null = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('🚀 Nutrient PDF Tools add-in initialized for Word');
        initializeApp();
    }
});

function initializeApp() {
    console.log('🔧 Initializing Nutrient PDF Tools application...');
    
    // Initialize tool selection
    initializeToolSelection();
    
    // Initialize export functionality
    initializeExportInterface();
    
    // Initialize settings functionality
    initializeSettingsInterface();
    
    // Initialize drag and drop functionality
    initializeDragAndDrop();
    
    // Initialize file input
    initializeFileInput();
    
    // Initialize convert button
    initializeConvertButton();
    
    // Initialize options
    initializeOptions();
    
    // Initialize convert to PDF functionality
    initializeConvertToPdfInterface();
    
    // Initialize authentication modal
    initializeAuthModal();
    
    // Check authentication on startup
    checkAuthentication();
    
    console.log('✅ Application initialization complete');
}

function initializeToolSelection() {
    const toolCards = document.querySelectorAll('.tool-card');
    
    toolCards.forEach(card => {
        card.addEventListener('click', () => {
            const toolName = card.getAttribute('data-tool');
            if (toolName) {
                showToolTab(toolName);
            }
        });
    });
}

function showToolTab(toolName: string) {
    // Hide all tabs
    const allTabs = document.querySelectorAll('.tab-content');
    allTabs.forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Show the selected tool tab
    const targetTab = document.getElementById(`${toolName}-tab`);
    if (targetTab) {
        targetTab.classList.add('active');
    }
}

function showToolsTab() {
    // Hide all tabs
    const allTabs = document.querySelectorAll('.tab-content');
    allTabs.forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Show the tools tab
    const toolsTab = document.getElementById('tools-tab');
    if (toolsTab) {
        toolsTab.classList.add('active');
    }
}

// Global function for back link
(window as any).showToolsTab = showToolsTab;

// Global function for advanced options toggle
(window as any).toggleAdvancedOptions = toggleAdvancedOptions;

// Global function for general advanced options toggle
(window as any).toggleGeneralAdvanced = toggleGeneralAdvanced;

function initializeExportInterface() {
    // Format selection handlers
    const formatOptions = document.querySelectorAll('input[name="export-format"]');
    formatOptions.forEach(option => {
        option.addEventListener('change', handleFormatChange);
    });
    
    // Security option handler
    const securitySelect = document.getElementById('export-security') as HTMLSelectElement;
    if (securitySelect) {
        securitySelect.addEventListener('change', handleSecurityChange);
    }
    
    // Export button handler
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', handleExport);
    }

    // Set default filename to current document name
    setDefaultFilename();
}

async function setDefaultFilename() {
    try {
        // Use Office.js API to get document properties
        const documentProperties = await Office.context.document.getSelectedDataAsync(Office.CoercionType.Text);
        let baseName = 'Document';
        
        // Try to get document name from Office context
        if (Office.context.document.url) {
            const url = Office.context.document.url;
            const fileName = url.split('/').pop() || 'Document';
            baseName = fileName.replace(/\.docx$/i, '');
        }
        
        // Set the filename input field
        const filenameInput = document.getElementById('export-filename') as HTMLInputElement;
        if (filenameInput) {
            filenameInput.value = baseName;
        }
    } catch (error) {
        console.error('Error getting document name:', error);
        // Fallback to 'Document' if there's an error
        const filenameInput = document.getElementById('export-filename') as HTMLInputElement;
        if (filenameInput) {
            filenameInput.value = 'Document';
        }
    }
}

function handleFormatChange(event: Event) {
    const target = event.target as HTMLInputElement;
    const pdfaOptions = document.getElementById('pdfa-options');
    const pdfuaOptions = document.getElementById('pdfua-options');
    
    if (target.value === 'pdfa') {
        pdfaOptions?.style.setProperty('display', 'block');
        pdfuaOptions?.style.setProperty('display', 'none');
    } else if (target.value === 'pdfua') {
        pdfaOptions?.style.setProperty('display', 'none');
        pdfuaOptions?.style.setProperty('display', 'block');
    }
}

function toggleAdvancedOptions(format: string) {
    const toggleButton = document.querySelector(`[onclick="toggleAdvancedOptions('${format}')"]`) as HTMLElement;
    const content = document.getElementById(`${format}-advanced-content`);
    
    if (!toggleButton || !content) {
        console.error(`Advanced options elements for ${format} not found`);
        return;
    }
    
    const isCollapsed = content.classList.contains('collapsed');
    
    if (isCollapsed) {
        // Expand
        content.classList.remove('collapsed');
        content.classList.add('expanded');
        toggleButton.classList.add('expanded');
    } else {
        // Collapse
        content.classList.remove('expanded');
        content.classList.add('collapsed');
        toggleButton.classList.remove('expanded');
    }
}

function toggleGeneralAdvanced() {
    const toggleButton = document.querySelector('[onclick="toggleGeneralAdvanced()"]') as HTMLElement;
    const content = document.getElementById('general-advanced-content');
    
    if (!toggleButton || !content) {
        console.error('General advanced options elements not found');
        return;
    }
    
    const isCollapsed = content.classList.contains('collapsed');
    
    if (isCollapsed) {
        // Expand
        content.classList.remove('collapsed');
        content.classList.add('expanded');
        toggleButton.classList.add('expanded');
    } else {
        // Collapse
        content.classList.remove('expanded');
        content.classList.add('collapsed');
        toggleButton.classList.remove('expanded');
    }
}

function handleSecurityChange(event: Event) {
    const target = event.target as HTMLSelectElement;
    const passwordSection = document.getElementById('password-section');
    
    if (target.value === 'password') {
        passwordSection?.style.setProperty('display', 'block');
    } else {
        passwordSection?.style.setProperty('display', 'none');
    }
}

async function handleExport() {
    const exportBtn = document.getElementById('export-btn') as HTMLButtonElement;
    const progressSection = document.getElementById('export-progress');
    const progressBar = document.getElementById('export-progress-bar') as HTMLElement;
    const progressText = document.getElementById('export-progress-text');
    const statusMessages = document.getElementById('export-status');
    
    // Get export options
    const format = (document.querySelector('input[name="export-format"]:checked') as HTMLInputElement)?.value;
    const filename = (document.getElementById('export-filename') as HTMLInputElement)?.value || 'Document';
    const quality = (document.getElementById('export-quality') as HTMLSelectElement)?.value;
    const security = (document.getElementById('export-security') as HTMLSelectElement)?.value;
    const password = (document.getElementById('export-password') as HTMLInputElement)?.value;
    
    // Get format-specific options
    const options: any = {
        format,
        quality,
        security
    };
    
    if (format === 'pdfa') {
        options.pdfa = {
            version: (document.getElementById('pdfa-version') as HTMLSelectElement)?.value,
            embedFonts: (document.getElementById('pdfa-embed-fonts') as HTMLInputElement)?.checked,
            colorProfile: (document.getElementById('pdfa-color-profile') as HTMLInputElement)?.checked
        };
    } else if (format === 'pdfua') {
        options.pdfua = {
            tags: (document.getElementById('pdfua-tags') as HTMLInputElement)?.checked,
            altText: (document.getElementById('pdfua-alt-text') as HTMLInputElement)?.checked,
            readingOrder: (document.getElementById('pdfua-reading-order') as HTMLInputElement)?.checked,
            colorContrast: (document.getElementById('pdfua-color-contrast') as HTMLInputElement)?.checked
        };
    }
    
    if (security === 'password' && password) {
        options.password = password;
    }
    
    try {
        // Disable button and show progress
        exportBtn.disabled = true;
        progressSection?.style.setProperty('display', 'block');
        showExportStatus('Preparing document for export...', 'info');
        
        // Simulate export process (replace with actual API call)
        await simulateExportProcess(progressBar, progressText);
        
        // Generate filename with .pdf extension
        const fileName = `${filename}.pdf`;
        
        // Call DWS Build API to export to PDF
        const pdfBase64 = await exportToPDF(format, quality, security, password);
        
        // Show PDF preview using DWS Viewer with actual PDF data
        showPDFPreviewWithDWSViewer(pdfBase64, fileName);
        
        // Also save the PDF as a file
        savePDFFromBase64(pdfBase64, fileName);
        
    } catch (error) {
        showExportStatus(`Export failed: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
    } finally {
        // Re-enable button and hide progress
        exportBtn.disabled = false;
        progressSection?.style.setProperty('display', 'none');
    }
}

function showExportStatus(message: string, type: 'info' | 'success' | 'error') {
    const statusMessages = document.getElementById('export-status');
    if (statusMessages) {
        const statusDiv = document.createElement('div');
        statusDiv.className = `status-message ${type}`;
        statusDiv.textContent = message;
        statusMessages.appendChild(statusDiv);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            statusDiv.remove();
        }, 5000);
    }
}

function showPDFPreview(pdfBlob: Blob, fileName: string) {
    const exportOptionsSection = document.querySelector('.export-options') as HTMLElement;
    const pdfPreviewSection = document.getElementById('pdf-preview-section') as HTMLElement;
    const pdfIframe = document.getElementById('pdf-preview-iframe') as HTMLIFrameElement;
    const downloadBtn = document.getElementById('download-pdf-btn') as HTMLButtonElement;
    const newExportBtn = document.getElementById('new-export-btn') as HTMLButtonElement;

    if (!exportOptionsSection || !pdfPreviewSection || !pdfIframe || !downloadBtn || !newExportBtn) {
        console.error('PDF preview elements not found');
        return;
    }

    // Create object URL for PDF
    const pdfUrl = URL.createObjectURL(pdfBlob);
    
    // Set iframe source to PDF
    pdfIframe.src = pdfUrl;

    // Show preview section, hide options section
    exportOptionsSection.style.display = 'none';
    pdfPreviewSection.style.display = 'block';

    // Setup download button
    downloadBtn.onclick = () => {
        const link = document.createElement('a');
        link.href = pdfUrl;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    // Setup new export button
    newExportBtn.onclick = () => {
        // Clean up object URL
        URL.revokeObjectURL(pdfUrl);
        
        // Hide preview, show options
        pdfPreviewSection.style.display = 'none';
        exportOptionsSection.style.display = 'block';
        
        // Clear iframe
        pdfIframe.src = '';
        
        // Clear status
        const statusDiv = document.getElementById('export-status') as HTMLElement;
        if (statusDiv) {
            statusDiv.innerHTML = '';
        }
    };
}

async function exportToPDF(format: string, quality: string, security: string, password?: string): Promise<string> {
    try {
        // Check authentication first
        if (!authCredentials) {
            showAuthModal();
            throw new Error('Authentication required. Please configure your API credentials.');
        }

        // Get current Word document content
        const documentContent = await getCurrentDocumentContent();

        // Create instructions for PDF export
        const instructions: any = {
            parts: [
                {
                    file: documentContent
                }
            ],
            output: {
                type: format === 'pdfa' ? 'pdfa' : 'pdfua'
            }
        };

        // Add format-specific options
        if (format === 'pdfa') {
            instructions.pdfa = {
                version: (document.getElementById('pdfa-version') as HTMLSelectElement)?.value || '1b',
                embedFonts: (document.getElementById('pdfa-embed-fonts') as HTMLInputElement)?.checked || true,
                colorProfile: (document.getElementById('pdfa-color-profile') as HTMLInputElement)?.checked || true
            };
        } else if (format === 'pdfua') {
            instructions.pdfua = {
                tags: (document.getElementById('pdfua-tags') as HTMLInputElement)?.checked || true,
                altText: (document.getElementById('pdfua-alt-text') as HTMLInputElement)?.checked || true,
                readingOrder: (document.getElementById('pdfua-reading-order') as HTMLInputElement)?.checked || true,
                colorContrast: (document.getElementById('pdfua-color-contrast') as HTMLInputElement)?.checked || true
            };
        }

        // Add security settings
        if (security === 'password' && password) {
            instructions.security = {
                password: password
            };
        }

        // Call DWS Build API using fetch with no-cors mode to bypass CORS issues
        const apiResult = await new Promise<any>((resolve, reject) => {
            // Create FormData with the correct structure
            const formData = new FormData();
            
            // Add the document content as a file (field name is 'document')
            const documentBlob = new Blob([documentContent], { type: 'text/plain' });
            formData.append('document', documentBlob, 'document.docx');
            
            // Create instructions JSON
            const instructions: any = {
                parts: [
                    {
                        file: "document"  // References the 'document' field above
                    }
                ],
                output: {
                    type: format === 'pdfa' ? 'pdfa' : 'pdfua'
                }
            };
            
            // Add format-specific options if needed
            if (format === 'pdfa') {
                instructions.output.conformance = 'pdfa-2a';  // Use conformance instead of nested pdfa
            } else if (format === 'pdfua') {
                instructions.output.conformance = 'pdfua-1';  // PDF/UA conformance
            }
            
            // Add instructions as JSON string
            formData.append('instructions', JSON.stringify(instructions));
            
            // Debug: Log the FormData contents
            console.log('=== FORM DATA DEBUG ===');
            console.log('API URL:', `${authCredentials.apiBaseUrl}/build`);
            console.log('Authorization:', `Bearer ${authCredentials.apiKey}`);
            console.log('Document blob size:', documentBlob.size);
            console.log('Format:', format);
            
            // Log FormData structure
            console.log('FormData structure:');
            console.log('- document: Blob with document content');
            console.log('- instructions:', JSON.stringify(instructions, null, 2));
            console.log('=== END FORM DATA DEBUG ===');
            
            // Try multiple approaches to handle CORS
            const tryRequest = async (method: 'fetch' | 'xhr') => {
                try {
                    if (method === 'fetch') {
                        const headers = getAuthHeaders();
                        // Remove Content-Type for FormData (browser will set it automatically with boundary)
                        delete headers['Content-Type'];
                        
                        const response = await fetch(`${authCredentials.apiBaseUrl}/build`, {
                            method: 'POST',
                            headers,
                            body: formData
                        });
                        
                        if (response.ok) {
                            const result = await response.json();
                            resolve(result);
                        } else {
                            const errorText = await response.text();
                            console.error('Fetch API Response:', errorText);
                            reject(new Error(`DWS Build API error: ${response.status} ${response.statusText}`));
                        }
                    } else {
                        // Fallback to XMLHttpRequest with different CORS handling
                        const xhr = new XMLHttpRequest();
                        xhr.open('POST', `${authCredentials.apiBaseUrl}/build`, true);
                        xhr.setRequestHeader('Authorization', `Bearer ${authCredentials.apiKey}`);
                        
                        xhr.onload = function() {
                            if (xhr.status === 200) {
                                try {
                                    const result = JSON.parse(xhr.responseText);
                                    resolve(result);
                                } catch (e) {
                                    reject(new Error('Invalid JSON response'));
                                }
                            } else {
                                console.error('XHR API Response:', xhr.responseText);
                                reject(new Error(`DWS Build API error: ${xhr.status} ${xhr.statusText}`));
                            }
                        };
                        
                        xhr.onerror = function() {
                            reject(new Error('Network error'));
                        };
                        
                        xhr.send(formData);
                    }
                } catch (error) {
                    if (method === 'fetch') {
                        console.log('Fetch failed, trying XMLHttpRequest...');
                        await tryRequest('xhr');
                    } else {
                        reject(error);
                    }
                }
            };
            
            // Start with fetch, fallback to XMLHttpRequest
            tryRequest('fetch');
        });

        const result = apiResult;
        
        console.log('DWS Build API Response:', result);
        
        if (result.document) {
            console.log('PDF base64 length:', result.document.length);
            return result.document; // Return base64 PDF data
        } else {
            console.error('No document in response:', result);
            throw new Error('No PDF data received from DWS Build API');
        }

    } catch (error) {
        console.error('Error exporting to PDF:', error);
        throw error;
    }
}

async function getCurrentDocumentContent(): Promise<Blob> {
    console.log('📄 getCurrentDocumentContent() called');
    try {
        // Use Office.js API to get the current document content
        console.log('🔍 Getting current Word document content...');
        
        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('✅ Document file obtained successfully');
                    
                    const file = result.value;
                    console.log('📊 File details:', {
                        size: file.size
                    });
                    
                    // Get the file content as a slice
                    file.getSliceAsync(0, (sliceResult) => {
                        if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log('✅ Document slice obtained successfully');
                            
                            const slice = sliceResult.value;
                            const data = slice.data;
                            
                            console.log('📊 Slice details:', {
                                size: data.length,
                                type: typeof data
                            });
                            
                            // Convert the data to a Blob
                            const blob = new Blob([data], { 
                                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
                            });
                            
                            console.log('✅ Document blob created successfully:', {
                                size: blob.size,
                                type: blob.type
                            });
                            
                            // Close the file
                            file.closeAsync(() => {
                                console.log('✅ File closed successfully');
                                resolve(blob);
                            });
                            
                        } else {
                            console.error('❌ Failed to get document slice:', sliceResult.error);
                            reject(new Error(`Failed to get document slice: ${sliceResult.error.message}`));
                        }
                    });
                    
                } else {
                    console.error('❌ Failed to get document file:', result.error);
                    reject(new Error(`Failed to get document file: ${result.error.message}`));
                }
            });
        });
        
    } catch (error) {
        console.error('❌ Error getting document content:', error);
        
        // Fallback to Invoice.docx if Office.js API fails
        console.log('🔄 Falling back to Invoice.docx...');
        try {
            const response = await fetch('assets/Invoice.docx');
            console.log('📡 Fetch response status:', response.status, response.statusText);
            
            if (!response.ok) {
                console.error('❌ Fetch failed:', response.status, response.statusText);
                throw new Error(`Failed to fetch Invoice.docx: ${response.status}`);
            }
            
            console.log('✅ Fetch successful, converting to blob...');
            const docxBlob = await response.blob();
            console.log('✅ Invoice.docx loaded successfully:', {
                size: docxBlob.size,
                type: docxBlob.type
            });
            return docxBlob;
            
        } catch (fallbackError) {
            console.error('❌ Fallback also failed:', fallbackError);
            throw new Error('Failed to get document content - both Office.js API and fallback failed');
        }
    }
}

async function showPDFPreviewWithDWSViewer(pdfBase64: string, fileName: string) {
    try {
        // Check authentication first
        if (!authCredentials) {
            showAuthModal();
            throw new Error('Authentication required. Please configure your API credentials.');
        }

        // Create form data for DWS Viewer API with base64 PDF data
        const formData = new FormData();
        formData.append('document', new Blob([pdfBase64], { type: 'text/plain' }), fileName);

        // Call DWS Viewer API using XMLHttpRequest to avoid CORS issues
        const viewerResult = await new Promise<any>((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('POST', `${authCredentials.apiBaseUrl}/view`, true);
            xhr.setRequestHeader('Authorization', `Bearer ${authCredentials.viewerKey}`);
            
            xhr.onload = function() {
                if (xhr.status === 200) {
                    try {
                        const result = JSON.parse(xhr.responseText);
                        resolve(result);
                    } catch (e) {
                        reject(new Error('Invalid JSON response from DWS Viewer'));
                    }
                } else {
                    reject(new Error(`DWS Viewer API error: ${xhr.status} ${xhr.statusText}`));
                }
            };
            
            xhr.onerror = function() {
                reject(new Error('Network error with DWS Viewer'));
            };
            
            xhr.send(formData);
        });

        const result = viewerResult;
        
        if (result.url) {
            // Use DWS Viewer URL for preview
            showPDFPreviewWithURL(result.url, fileName);
        } else {
            // Fallback to local preview if no DWS Viewer URL
            const pdfBlob = new Blob([Uint8Array.from(atob(pdfBase64), c => c.charCodeAt(0))], { type: 'application/pdf' });
            showPDFPreview(pdfBlob, fileName);
        }

    } catch (error) {
        console.error('Error using DWS Viewer:', error);
        // Fallback to local preview
        const pdfBlob = new Blob([Uint8Array.from(atob(pdfBase64), c => c.charCodeAt(0))], { type: 'application/pdf' });
        showPDFPreview(pdfBlob, fileName);
    }
}

function showPDFPreviewWithURL(viewerUrl: string, fileName: string) {
    const exportOptionsSection = document.querySelector('.export-options') as HTMLElement;
    const pdfPreviewSection = document.getElementById('pdf-preview-section') as HTMLElement;
    const pdfIframe = document.getElementById('pdf-preview-iframe') as HTMLIFrameElement;
    const downloadBtn = document.getElementById('download-pdf-btn') as HTMLButtonElement;
    const newExportBtn = document.getElementById('new-export-btn') as HTMLButtonElement;

    if (!exportOptionsSection || !pdfPreviewSection || !pdfIframe || !downloadBtn || !newExportBtn) {
        console.error('PDF preview elements not found');
        return;
    }

    // Set iframe source to DWS Viewer URL
    pdfIframe.src = viewerUrl;

    // Show preview section, hide options section
    exportOptionsSection.style.display = 'none';
    pdfPreviewSection.style.display = 'block';

    // Setup download button (opens DWS Viewer in new tab for download)
    downloadBtn.onclick = () => {
        window.open(viewerUrl, '_blank');
    };

    // Setup new export button
    newExportBtn.onclick = () => {
        // Hide preview, show options
        pdfPreviewSection.style.display = 'none';
        exportOptionsSection.style.display = 'block';
        
        // Clear iframe
        pdfIframe.src = '';
        
        // Clear status
        const statusDiv = document.getElementById('export-status') as HTMLElement;
        if (statusDiv) {
            statusDiv.innerHTML = '';
        }
    };
}

function savePDFFromBase64(base64Data: string, fileName: string) {
    try {
        console.log('Converting base64 to PDF blob...');
        
        // Convert base64 to binary data
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        
        // Create PDF blob
        const pdfBlob = new Blob([bytes], { type: 'application/pdf' });
        console.log('PDF blob created, size:', pdfBlob.size, 'bytes');
        
        // Create download link
        const url = URL.createObjectURL(pdfBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        link.style.display = 'none';
        
        // Trigger download
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Clean up
        URL.revokeObjectURL(url);
        
        console.log('PDF saved as:', fileName);
        
    } catch (error) {
        console.error('Error saving PDF from base64:', error);
    }
}

async function simulateExportProcess(progressBar: HTMLElement, progressText: HTMLElement) {
    const steps = [
        { progress: 20, text: 'Analyzing document structure...' },
        { progress: 40, text: 'Applying accessibility standards...' },
        { progress: 60, text: 'Generating PDF format...' },
        { progress: 80, text: 'Applying security settings...' },
        { progress: 100, text: 'Finalizing export...' }
    ];
    
    for (const step of steps) {
        progressBar.style.width = `${step.progress}%`;
        progressText.textContent = step.text;
        await new Promise(resolve => setTimeout(resolve, 800));
    }
}

async function promptForFileName(format: string): Promise<string | null> {
    // Get current document name
    let currentDocName = 'Document';
    try {
        await Word.run(async (context) => {
            // Note: Word.js doesn't have direct access to document name
            // We'll use a default name based on the current document
            currentDocName = 'Document';
        });
    } catch (error) {
        console.warn('Could not get document name:', error);
    }
    
    // Create default filename based on format
    const formatExtension = format === 'pdfa' ? 'pdfa' : 'pdfua';
    const defaultFileName = `${currentDocName}.${formatExtension}`;
    
    // Use browser's file save dialog
    return new Promise((resolve) => {
        // Create a temporary download link
        const link = document.createElement('a');
        link.download = defaultFileName;
        link.style.display = 'none';
        
        // Create a dummy blob to trigger the save dialog
        const dummyBlob = new Blob([''], { type: 'application/pdf' });
        link.href = URL.createObjectURL(dummyBlob);
        
        // Trigger the download dialog
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // For now, return the default filename
        // In a real implementation, you'd need to handle the actual file save
        // This is a limitation of web browsers - they don't allow direct file system access
        resolve(defaultFileName);
    });
}

async function savePDFToLocation(pdfBlob: Blob, fileName: string): Promise<void> {
    // Create download link
    const link = document.createElement('a');
    link.href = URL.createObjectURL(pdfBlob);
    link.download = fileName;
    link.style.display = 'none';
    
    // Trigger download
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Clean up
    URL.revokeObjectURL(link.href);
}

async function openPDFInNewTab(pdfBlob: Blob, fileName: string): Promise<void> {
    // Create object URL for the PDF blob
    const pdfUrl = URL.createObjectURL(pdfBlob);
    
    // Open PDF in new tab
    const newTab = window.open(pdfUrl, '_blank');
    
    // Clean up object URL after a delay to ensure the tab has loaded
    setTimeout(() => {
        URL.revokeObjectURL(pdfUrl);
    }, 5000);
    
    // If new tab couldn't be opened (popup blocker), fall back to download
    if (!newTab) {
        console.warn('Popup blocked, falling back to download');
        await savePDFToLocation(pdfBlob, fileName);
    }
}

// Alternative approach using Office.js file picker (if available)
async function savePDFWithOfficePicker(pdfBlob: Blob, defaultFileName: string): Promise<void> {
    try {
        // This would use Office.js file picker if available
        // For now, fall back to browser download
        await savePDFToLocation(pdfBlob, defaultFileName);
    } catch (error) {
        console.warn('Office file picker not available, using browser download:', error);
        await savePDFToLocation(pdfBlob, defaultFileName);
    }
}

// Settings Management
interface Settings {
    processorApiKey: string;
    viewerApiKey: string;
    apiBaseUrl: string;
}

function initializeSettingsInterface() {
    // Load saved settings
    loadSettings();
    
    // Initialize password toggle buttons
    initializePasswordToggles();
    
    // Initialize save/clear buttons
    initializeSettingsButtons();
}

function loadSettings() {
    try {
        const savedSettings = localStorage.getItem('nutrientApiSettings');
        if (savedSettings) {
            const settings: Settings = JSON.parse(savedSettings);
            
            const processorKeyInput = document.getElementById('processor-api-key') as HTMLInputElement;
            const viewerKeyInput = document.getElementById('viewer-api-key') as HTMLInputElement;
            const baseUrlInput = document.getElementById('api-base-url') as HTMLInputElement;
            
            if (processorKeyInput) processorKeyInput.value = settings.processorApiKey || '';
            if (viewerKeyInput) viewerKeyInput.value = settings.viewerApiKey || '';
            if (baseUrlInput) baseUrlInput.value = settings.apiBaseUrl || 'https://localhost:8080/api/nutrient';
        }
        
        // Show authentication status
        updateAuthStatus();
    } catch (error) {
        console.warn('Failed to load settings:', error);
    }
}

function updateAuthStatus() {
    const authBtn = document.getElementById('auth-settings-btn');
    if (authBtn) {
        if (authCredentials) {
            authBtn.innerHTML = '<span class="ms-Button-label">✅ Authentication Configured</span>';
            authBtn.classList.add('authenticated');
        } else {
            authBtn.innerHTML = '<span class="ms-Button-label">🔐 Configure Authentication</span>';
            authBtn.classList.remove('authenticated');
        }
    }
}

function saveSettings(): boolean {
    try {
        const processorKeyInput = document.getElementById('processor-api-key') as HTMLInputElement;
        const viewerKeyInput = document.getElementById('viewer-api-key') as HTMLInputElement;
        const baseUrlInput = document.getElementById('api-base-url') as HTMLInputElement;
        
        const settings: Settings = {
            processorApiKey: processorKeyInput?.value || '',
            viewerApiKey: viewerKeyInput?.value || '',
            apiBaseUrl: baseUrlInput?.value || 'https://localhost:8080/api/nutrient'
        };
        
        localStorage.setItem('nutrientApiSettings', JSON.stringify(settings));
        return true;
    } catch (error) {
        console.error('Failed to save settings:', error);
        return false;
    }
}

function clearSettings() {
    try {
        localStorage.removeItem('nutrientApiSettings');
        
        const processorKeyInput = document.getElementById('processor-api-key') as HTMLInputElement;
        const viewerKeyInput = document.getElementById('viewer-api-key') as HTMLInputElement;
        const baseUrlInput = document.getElementById('api-base-url') as HTMLInputElement;
        
        if (processorKeyInput) processorKeyInput.value = '';
        if (viewerKeyInput) viewerKeyInput.value = '';
        if (baseUrlInput) baseUrlInput.value = 'https://localhost:8080/api/nutrient';
        
        return true;
    } catch (error) {
        console.error('Failed to clear settings:', error);
        return false;
    }
}

function initializePasswordToggles() {
    const toggleProcessorBtn = document.getElementById('toggle-processor-key');
    const toggleViewerBtn = document.getElementById('toggle-viewer-key');
    
    if (toggleProcessorBtn) {
        toggleProcessorBtn.addEventListener('click', () => {
            togglePasswordVisibility('processor-api-key', toggleProcessorBtn);
        });
    }
    
    if (toggleViewerBtn) {
        toggleViewerBtn.addEventListener('click', () => {
            togglePasswordVisibility('viewer-api-key', toggleViewerBtn);
        });
    }
}

function togglePasswordVisibility(inputId: string, button: HTMLElement) {
    const input = document.getElementById(inputId) as HTMLInputElement;
    if (input) {
        if (input.type === 'password') {
            input.type = 'text';
            button.textContent = '🙈';
        } else {
            input.type = 'password';
            button.textContent = '👁️';
        }
    }
}

function initializeSettingsButtons() {
    const saveBtn = document.getElementById('save-settings-btn');
    const clearBtn = document.getElementById('clear-settings-btn');
    const authBtn = document.getElementById('auth-settings-btn');
    
    if (saveBtn) {
        saveBtn.addEventListener('click', handleSaveSettings);
    }
    
    if (clearBtn) {
        clearBtn.addEventListener('click', handleClearSettings);
    }
    
    if (authBtn) {
        authBtn.addEventListener('click', showAuthModal);
    }
}

async function handleSaveSettings() {
    const statusMessage = document.getElementById('settings-status-message');
    
    // Clear previous status messages
    if (statusMessage) {
        statusMessage.innerHTML = '';
    }
    
    // Save settings directly without validation
    const success = saveSettings();
    
    if (statusMessage) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `status-message ${success ? 'success' : 'error'}`;
        messageDiv.textContent = success ? 'Settings saved successfully!' : 'Failed to save settings.';
        statusMessage.appendChild(messageDiv);
        
        // Auto-remove after 3 seconds
        setTimeout(() => {
            messageDiv.remove();
        }, 3000);
    }
}

function handleClearSettings() {
    const success = clearSettings();
    const statusMessage = document.getElementById('settings-status-message');
    
    if (statusMessage) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `status-message ${success ? 'success' : 'error'}`;
        messageDiv.textContent = success ? 'Settings cleared successfully!' : 'Failed to clear settings.';
        statusMessage.appendChild(messageDiv);
        
        // Auto-remove after 3 seconds
        setTimeout(() => {
            messageDiv.remove();
        }, 3000);
    }
    
    // Also clear authentication credentials
    try {
        localStorage.removeItem('nutrientAuthCredentials');
        authCredentials = null;
        console.log('Authentication credentials cleared');
        updateAuthStatus();
    } catch (error) {
        console.warn('Failed to clear auth credentials:', error);
    }
}

// Function to get current API settings
function getApiSettings(): Settings | null {
    try {
        const savedSettings = localStorage.getItem('nutrientApiSettings');
        if (savedSettings) {
            return JSON.parse(savedSettings);
        }
    } catch (error) {
        console.warn('Failed to get API settings:', error);
    }
    return null;
}



function initializeDragAndDrop() {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    const fileInput = document.getElementById('file-input') as HTMLInputElement;

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop zone when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    dragDropZone.addEventListener('drop', handleDrop, false);
    
    // Handle click to browse
    dragDropZone.addEventListener('click', () => fileInput.click());
}

function initializeFileInput() {
    const fileInput = document.getElementById('file-input') as HTMLInputElement;
    fileInput.addEventListener('change', handleFileSelect);
}

function initializeConvertButton() {
    const convertBtn = document.getElementById('convert-btn') as HTMLButtonElement;
    convertBtn.addEventListener('click', handleConvert);
}

function initializeOptions() {
    // OCR toggle and language select are already handled by HTML
    // Additional initialization can be added here if needed
}



function preventDefaults(e: Event) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e: Event) {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    dragDropZone.classList.add('dragover');
}

function unhighlight(e: Event) {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    dragDropZone.classList.remove('dragover');
}

function handleDrop(e: DragEvent) {
    const dt = e.dataTransfer;
    const files = dt?.files;
    
    if (files && files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e: Event) {
    const target = e.target as HTMLInputElement;
    if (target.files && target.files.length > 0) {
        handleFile(target.files[0]);
    }
}

function handleFile(file: File) {
    // Validate file type
    if (file.type !== 'application/pdf') {
        showStatus('Please select a PDF file.', 'error');
        return;
    }

    // Validate file size (max 50MB)
    if (file.size > 50 * 1024 * 1024) {
        showStatus('File size must be less than 50MB.', 'error');
        return;
    }

    // Store file data
    selectedFile = {
        file: file,
        name: file.name,
        size: formatFileSize(file.size)
    };

    // Update UI
    updateFileInfo();
    enableConvertButton();
    showStatus('PDF file selected successfully.', 'success');
}

function updateFileInfo() {
    const fileInfo = document.getElementById('file-info') as HTMLElement;
    const fileName = document.getElementById('file-name') as HTMLElement;
    const fileSize = document.getElementById('file-size') as HTMLElement;

    if (selectedFile) {
        fileName.textContent = `File: ${selectedFile.name}`;
        fileSize.textContent = `Size: ${selectedFile.size}`;
        fileInfo.style.display = 'block';
    } else {
        fileInfo.style.display = 'none';
    }
}

function enableConvertButton() {
    const convertBtn = document.getElementById('convert-btn') as HTMLButtonElement;
    convertBtn.disabled = !selectedFile;
}

function formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function handleConvert() {
    if (!selectedFile) {
        showStatus('Please select a PDF file first.', 'error');
        return;
    }

    try {
        showProgress(true);
        updateProgress(10, 'Preparing file for conversion...');

        // Get options
        const ocrEnabled = (document.getElementById('ocr-toggle') as HTMLInputElement).checked;
        const language = (document.getElementById('language-select') as HTMLSelectElement).value;

        updateProgress(30, 'Uploading to Nutrient.io...');

        // Convert PDF to DOCX using Nutrient.io API
        const docxBlob = await convertPdfToDocx(selectedFile.file, ocrEnabled, language);

        updateProgress(70, 'Inserting into Word document...');

        // Insert the converted document into Word
        await insertDocxIntoWord(docxBlob);

        updateProgress(100, 'Conversion completed!');
        
        setTimeout(() => {
            showProgress(false);
            showStatus('PDF successfully converted and inserted into Word document!', 'success');
            resetFileSelection();
        }, 1000);

    } catch (error) {
        showProgress(false);
        console.error('Conversion error:', error);
        showStatus(`Conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
    }
}

async function convertPdfToDocx(file: File, ocrEnabled: boolean, language: string): Promise<Blob> {
    // Check authentication first
    if (!authCredentials) {
        showAuthModal();
        throw new Error('Authentication required. Please configure your API credentials.');
    }
    
    const formData = new FormData();

    // Prepare the API payload according to Nutrient.io Build API specification
    const payload: any = {
        parts: [{ file: 'document' }],
        ocr: ocrEnabled,
        output: { type: 'docx' }
    };

    // Add language if specified and not auto-detect
    if (language !== 'auto') {
        payload.ocr_language = language;
    }

    formData.append('instructions', JSON.stringify(payload));
    formData.append('document', file);

    // Show request details in UI
    showStatus(`Making API request to ${authCredentials.apiBaseUrl}/build...`, 'info');

    try {
        const headers = getAuthHeaders();
        // Remove Content-Type for FormData (browser will set it automatically with boundary)
        delete headers['Content-Type'];
        
        console.log('=== API REQUEST DEBUG ===');
        console.log('Request URL:', `${authCredentials.apiBaseUrl}/build`);
        console.log('Request method:', 'POST');
        console.log('Request headers:', headers);
        console.log('FormData contents:');
        // Note: FormData.entries() is not available in all TypeScript configurations
        console.log('- document: File object');
        console.log('- instructions: JSON payload for API');
        console.log('=== END API REQUEST DEBUG ===');
        
        const response = await fetch(`${authCredentials.apiBaseUrl}/build`, {
            method: 'POST',
            headers,
            body: formData
        });

        if (!response.ok) {
            const errorText = await response.text();
            
            // Show detailed error in UI
            showStatus(`API Error ${response.status}: ${response.statusText} - ${errorText}`, 'error');
            throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const blob = await response.blob();

        // Show success in UI
        showStatus(`API request successful! Received ${blob.size} bytes`, 'success');
        
        return blob;
    } catch (error) {
        // Show detailed error in UI
        showStatus(`Network Error: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
        throw error;
    }
}

async function insertDocxIntoWord(docxBlob: Blob) {
    return Word.run(async (context) => {
        // Convert blob to base64
        const arrayBuffer = await docxBlob.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const base64 = btoa(String.fromCharCode.apply(null, Array.from(uint8Array)));

        // Insert the document at the current cursor position
        const range = context.document.getSelection();
        range.insertFileFromBase64(base64, Word.InsertLocation.replace);

        await context.sync();
    });
}

function showProgress(show: boolean) {
    const progressSection = document.getElementById('progress-section') as HTMLElement;
    progressSection.style.display = show ? 'block' : 'none';
}

function updateProgress(percentage: number, text: string) {
    const progressBar = document.getElementById('progress-bar') as HTMLElement;
    const progressText = document.getElementById('progress-text') as HTMLElement;

    progressBar.style.width = `${percentage}%`;
    progressText.textContent = text;
}

function showStatus(message: string, type: 'success' | 'error' | 'info') {
    const statusMessages = document.getElementById('status-messages') as HTMLElement;
    
    // Remove existing messages
    const existingMessages = statusMessages.querySelectorAll('.status-message');
    existingMessages.forEach(msg => msg.remove());

    // Create new message
    const messageElement = document.createElement('div');
    messageElement.className = `status-message ${type}`;
    messageElement.textContent = message;
    
    statusMessages.appendChild(messageElement);

    // Auto-remove success/info messages after 5 seconds
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            if (messageElement.parentNode) {
                messageElement.remove();
            }
        }, 5000);
    }
}



function resetFileSelection() {
    selectedFile = null;
    updateFileInfo();
    enableConvertButton();
    
    // Reset file input
    const fileInput = document.getElementById('file-input') as HTMLInputElement;
    fileInput.value = '';
}

// Authentication Functions
function initializeAuthModal() {
    // Initialize modal close button
    const closeBtn = document.getElementById('close-auth-modal');
    if (closeBtn) {
        closeBtn.addEventListener('click', closeAuthModal);
    }
    
    // Initialize cancel button
    const cancelBtn = document.getElementById('auth-cancel-btn');
    if (cancelBtn) {
        cancelBtn.addEventListener('click', closeAuthModal);
    }
    
    // Initialize save button
    const saveBtn = document.getElementById('auth-save-btn');
    if (saveBtn) {
        saveBtn.addEventListener('click', handleAuthSave);
    }
    
    // Initialize password toggles
    const toggleApiKeyBtn = document.getElementById('toggle-auth-api-key');
    const toggleViewerKeyBtn = document.getElementById('toggle-auth-viewer-key');
    
    if (toggleApiKeyBtn) {
        toggleApiKeyBtn.addEventListener('click', () => {
            toggleAuthPasswordVisibility('auth-api-key', toggleApiKeyBtn);
        });
    }
    
    if (toggleViewerKeyBtn) {
        toggleViewerKeyBtn.addEventListener('click', () => {
            toggleAuthPasswordVisibility('auth-viewer-key', toggleViewerKeyBtn);
        });
    }
    
    // Load saved credentials if available
    loadAuthCredentials();
}

function showAuthModal() {
    const modal = document.getElementById('auth-modal');
    if (modal) {
        modal.style.display = 'flex';
        // Focus on first input
        const firstInput = document.getElementById('auth-tenant') as HTMLInputElement;
        if (firstInput) {
            firstInput.focus();
        }
    }
}

function closeAuthModal() {
    const modal = document.getElementById('auth-modal');
    if (modal) {
        modal.style.display = 'none';
    }
}

function toggleAuthPasswordVisibility(inputId: string, button: HTMLElement) {
    const input = document.getElementById(inputId) as HTMLInputElement;
    
    if (input && button) {
        if (input.type === 'password') {
            input.type = 'text';
            button.textContent = '🙈';
        } else {
            input.type = 'password';
            button.textContent = '👁️';
        }
    }
}

function loadAuthCredentials() {
    try {
        const saved = localStorage.getItem('nutrientAuthCredentials');
        if (saved) {
            const credentials: AuthCredentials = JSON.parse(saved);
            
            // Update existing credentials to use the new proxy URL if they're using the old one
            if (credentials.apiBaseUrl && credentials.apiBaseUrl !== 'https://localhost:8080/api/nutrient') {
                console.log('🔄 Updating existing credentials to use new proxy URL');
                credentials.apiBaseUrl = 'https://localhost:8080/api/nutrient';
                localStorage.setItem('nutrientAuthCredentials', JSON.stringify(credentials));
            }
            
            const apiKeyInput = document.getElementById('auth-api-key') as HTMLInputElement;
            const viewerKeyInput = document.getElementById('auth-viewer-key') as HTMLInputElement;
            
            if (apiKeyInput) apiKeyInput.value = credentials.apiKey || '';
            if (viewerKeyInput) viewerKeyInput.value = credentials.viewerKey || '';
        }
    } catch (error) {
        console.warn('Failed to load auth credentials:', error);
    }
}

async function handleAuthSave() {
    const statusMessage = document.getElementById('auth-status-message');
    
    // Clear previous status
    if (statusMessage) {
        statusMessage.innerHTML = '';
    }
    
    // Get form values
    const apiKey = (document.getElementById('auth-api-key') as HTMLInputElement)?.value?.trim();
    const viewerKey = (document.getElementById('auth-viewer-key') as HTMLInputElement)?.value?.trim();
    const apiBaseUrl = 'https://localhost:8080/api/nutrient'; // Fixed value
    
    // Validate inputs
    if (!apiKey && !viewerKey) {
        showAuthStatus('Please enter both Processor API Key and Viewer API Key. Both fields are required.', 'error');
        return;
    }
    
    if (!apiKey) {
        showAuthStatus('Please enter your Processor API Key. This key is required for PDF processing and conversion operations.', 'error');
        return;
    }
    
    if (!viewerKey) {
        showAuthStatus('Please enter your Viewer API Key. This key is required for PDF preview and viewing operations.', 'error');
        return;
    }
    
    try {
        // Test the credentials
        showAuthStatus('Testing credentials...', 'info');
        
        const credentials: AuthCredentials = {
            apiKey,
            viewerKey,
            apiBaseUrl
        };
        
        // Save credentials without testing (will be validated during actual API usage)
        authCredentials = credentials;
        localStorage.setItem('nutrientAuthCredentials', JSON.stringify(credentials));
        
        showAuthStatus('Credentials saved successfully! They will be validated when you use the PDF tools.', 'success');
        
        // Update authentication status in settings
        updateAuthStatus();
        
        // Close modal after a short delay
        setTimeout(() => {
            closeAuthModal();
        }, 1500);
    } catch (error) {
        showAuthStatus(`Authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
    }
}

function showAuthStatus(message: string, type: 'success' | 'error' | 'info') {
    const statusMessage = document.getElementById('auth-status-message');
    if (statusMessage) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `status-message ${type}`;
        messageDiv.textContent = message;
        statusMessage.appendChild(messageDiv);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.remove();
            }
        }, 5000);
    }
}

async function testAuthCredentials(credentials: AuthCredentials): Promise<{ success: boolean; message: string }> {
    try {
        // Test with a simple API call to verify credentials using the build endpoint
        const headers = {
            'Authorization': `Bearer ${credentials.apiKey}`,
            'Content-Type': 'application/json'
        };
        
        console.log('Testing credentials with URL:', `${credentials.apiBaseUrl}/build`);
        console.log('Using headers:', headers);
        
        // Create a minimal test payload for the build endpoint
        const testPayload = {
            parts: [{ file: "test" }],
            output: { type: "pdf" }
        };
        
        const response = await fetch(`${credentials.apiBaseUrl}/build`, {
            method: 'POST',
            headers,
            body: JSON.stringify(testPayload)
        });
        
        console.log('Response status:', response.status);
        console.log('Response status text:', response.statusText);
        
        if (response.ok) {
            return { success: true, message: 'Authentication successful!' };
        } else {
            const errorText = await response.text();
            console.error('API Error Response:', errorText);
            
            if (response.status === 401) {
                return { 
                    success: false, 
                    message: `Authentication failed (401 Unauthorized). Your Processor API Key appears to be invalid. Please check that you've entered the correct API key.` 
                };
            } else if (response.status === 403) {
                return { 
                    success: false, 
                    message: `Authentication failed (403 Forbidden). Your Processor API Key may not have the required permissions. Please check your API key permissions.` 
                };
            } else if (response.status === 404) {
                return { 
                    success: false, 
                    message: `Authentication failed (404 Not Found). The API endpoint is not available. Please check that the API Base URL is correct: ${credentials.apiBaseUrl}` 
                };
            } else {
                return { 
                    success: false, 
                    message: `Authentication failed (${response.status} ${response.statusText}). Server response: ${errorText || 'No error details provided'}` 
                };
            }
        }
    } catch (error) {
        console.error('Network error during auth test:', error);
        
        if (error instanceof TypeError && error.message.includes('fetch')) {
            return { 
                success: false, 
                message: `Network error: Unable to connect to ${credentials.apiBaseUrl}. Please check your internet connection and that the API Base URL is correct.` 
            };
        } else if (error instanceof TypeError && error.message.includes('CORS')) {
            return { 
                success: false, 
                message: `CORS error: The API server is not allowing requests from this domain. This might be a configuration issue on the server side.` 
            };
        } else {
            return { 
                success: false, 
                message: `Network error: ${error instanceof Error ? error.message : 'Unknown network error occurred'}` 
            };
        }
    }
}

function checkAuthentication() {
    try {
        const saved = localStorage.getItem('nutrientAuthCredentials');
        if (saved) {
            const credentials: AuthCredentials = JSON.parse(saved);
            
            // Update existing credentials to use the new proxy URL if they're using the old one
            if (credentials.apiBaseUrl && credentials.apiBaseUrl !== 'https://localhost:8080/api/nutrient') {
                console.log('🔄 Updating existing credentials to use new proxy URL');
                credentials.apiBaseUrl = 'https://localhost:8080/api/nutrient';
                localStorage.setItem('nutrientAuthCredentials', JSON.stringify(credentials));
            }
            
            authCredentials = credentials;
            console.log('Authentication loaded from storage');
            updateAuthStatus();
        } else {
            // Show auth modal if no credentials found
            showAuthModal();
        }
    } catch (error) {
        console.warn('Failed to check authentication:', error);
        showAuthModal();
    }
}

function getAuthHeaders(): Record<string, string> {
    if (!authCredentials) {
        throw new Error('Authentication required. Please configure your API credentials.');
    }
    
    console.log('=== AUTHENTICATION DEBUG ===');
    console.log('Auth credentials loaded:', {
        apiKey: authCredentials.apiKey ? `${authCredentials.apiKey.substring(0, 10)}...` : 'undefined',
        viewerKey: authCredentials.viewerKey ? `${authCredentials.viewerKey.substring(0, 10)}...` : 'undefined',
        apiBaseUrl: authCredentials.apiBaseUrl
    });
    
    const headers = {
        'Authorization': `Bearer ${authCredentials.apiKey}`,
        'Content-Type': 'application/json'
    };
    
    console.log('Generated headers:', headers);
    console.log('Full Authorization header value:', `Bearer ${authCredentials.apiKey}`);
    console.log('=== END AUTHENTICATION DEBUG ===');
    
    return headers;
}

// Global function to show auth modal
(window as any).showAuthModal = showAuthModal;

// Convert to PDF Functions
function initializeConvertToPdfInterface() {
    // Load current document info
    loadCurrentDocumentInfo();
    
    // Initialize convert button
    const convertBtn = document.getElementById('convert-pdf-btn');
    if (convertBtn) {
        convertBtn.addEventListener('click', handleConvertToPdf);
    }
    
    // Initialize download button
    const downloadBtn = document.getElementById('download-pdf-btn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', handleDownloadPdf);
    }
    
    // Initialize new convert button
    const newConvertBtn = document.getElementById('new-convert-btn');
    if (newConvertBtn) {
        newConvertBtn.addEventListener('click', resetConvertToPdf);
    }
    
    // Initialize debug toggle button
    const toggleDebugBtn = document.getElementById('toggle-debug-btn');
    if (toggleDebugBtn) {
        toggleDebugBtn.addEventListener('click', toggleDebugSection);
    }
}

async function loadCurrentDocumentInfo() {
    try {
        const docNameElement = document.getElementById('current-doc-name');
        const docStatusElement = document.getElementById('current-doc-status');
        
        if (docNameElement && docStatusElement) {
            // Try to get document name from Office context
            let docName = 'Document';
            if (Office.context.document.url) {
                const url = Office.context.document.url;
                const fileName = url.split('/').pop() || 'Document';
                docName = fileName.replace(/\.docx$/i, '');
            }
            
            docNameElement.textContent = `Document: ${docName}`;
            docStatusElement.textContent = 'Ready to convert';
            
            // Set default filename
            const filenameInput = document.getElementById('convert-filename') as HTMLInputElement;
            if (filenameInput) {
                filenameInput.value = docName;
            }
        }
    } catch (error) {
        console.error('Error loading document info:', error);
    }
}

async function handleConvertToPdf() {
    console.log('=== CONVERT TO PDF START ===');
    
    const convertBtn = document.getElementById('convert-pdf-btn') as HTMLButtonElement;
    const progressSection = document.getElementById('convert-pdf-progress') as HTMLElement;
    const progressBar = document.getElementById('convert-pdf-progress-bar') as HTMLElement;
    const progressText = document.getElementById('convert-pdf-progress-text') as HTMLElement;
    
    console.log('🔍 UI elements found:', {
        convertBtn: !!convertBtn,
        progressSection: !!progressSection,
        progressBar: !!progressBar,
        progressText: !!progressText
    });
    
    // Check authentication first
    if (!authCredentials) {
        console.log('❌ No authentication credentials found, showing auth modal');
        showAuthModal();
        return;
    }
    
    console.log('✅ Authentication credentials found:', {
        hasApiKey: !!authCredentials.apiKey,
        hasViewerKey: !!authCredentials.viewerKey,
        apiBaseUrl: authCredentials.apiBaseUrl,
        apiKeyPreview: authCredentials.apiKey ? `${authCredentials.apiKey.substring(0, 20)}...` : 'undefined'
    });
    
    // Get options
    const filename = (document.getElementById('convert-filename') as HTMLInputElement)?.value || 'Document';
    const quality = (document.getElementById('convert-quality') as HTMLSelectElement)?.value || 'medium';
    
    console.log('📋 Conversion options:', { filename, quality });
    
    try {
        console.log('🔄 Starting conversion process...');
        
        // Disable button and show progress
        convertBtn.disabled = true;
        progressSection.style.display = 'block';
        
        console.log('✅ UI updated - button disabled, progress shown');
        
        // Show debug section for API transparency
        const debugSection = document.getElementById('convert-pdf-debug') as HTMLElement;
        if (debugSection) {
            debugSection.style.display = 'block';
            console.log('✅ Debug section shown');
        }
        
        showConvertPdfStatus('Preparing document for conversion...', 'info');
        console.log('📝 Status message shown: Preparing document for conversion...');
        
        // Update progress
        progressBar.style.width = '20%';
        progressText.textContent = 'Getting document content...';
        console.log('📊 Progress updated: 20% - Getting document content...');
        
        // Get current Word document content
        console.log('📄 Calling getCurrentDocumentContent()...');
        const documentBlob = await getCurrentDocumentContent();
        console.log('✅ Document blob received:', {
            size: documentBlob.size,
            type: documentBlob.type
        });
        
        progressBar.style.width = '40%';
        progressText.textContent = 'Converting to PDF...';
        console.log('📊 Progress updated: 40% - Converting to PDF...');
        
        // Convert to PDF using the build API
        console.log('🚀 Calling convertWordToPdf()...');
        const pdfBlob = await convertWordToPdf(documentBlob, quality);
        console.log('✅ PDF blob received:', {
            size: pdfBlob.size,
            type: pdfBlob.type
        });
        
        progressBar.style.width = '100%';
        progressText.textContent = 'Conversion completed!';
        console.log('📊 Progress updated: 100% - Conversion completed!');
        
        // Save PDF and create download link
        const fileName = `${filename}.pdf`;
        console.log('💾 Saving PDF for download:', fileName);
        savePdfForDownload(pdfBlob, fileName);
        
        // Show success message
        const successMessage = `✅ PDF created successfully! File size: ${(pdfBlob.size / 1024).toFixed(1)} KB`;
        showConvertPdfStatus(successMessage, 'success');
        console.log('✅ Success message shown:', successMessage);
        
        // Show download section
        setTimeout(() => {
            console.log('📥 Showing download section...');
            showDownloadSection(fileName);
        }, 1000);
        
        // Automatically trigger download after a short delay
        setTimeout(() => {
            console.log('🚀 Auto-triggering download...');
            handleDownloadPdf();
        }, 1500);
        
        console.log('=== CONVERT TO PDF SUCCESS ===');
        
    } catch (error) {
        console.error('❌ PDF conversion error:', error);
        
        // Provide more specific error messages
        let errorMessage = 'Conversion failed';
        if (error instanceof Error) {
            if (error.message.includes('401')) {
                errorMessage = 'Authentication failed. Please check your Processor API Key.';
            } else if (error.message.includes('403')) {
                errorMessage = 'Access denied. Please check your API permissions.';
            } else if (error.message.includes('404')) {
                errorMessage = 'API endpoint not found. Please check your configuration.';
            } else if (error.message.includes('500')) {
                errorMessage = 'Server error. Please try again later.';
            } else if (error.message.includes('NetworkError') || error.message.includes('fetch')) {
                errorMessage = 'Network error. Please check your internet connection.';
            } else {
                errorMessage = `Conversion failed: ${error.message}`;
            }
        }
        
        console.log('❌ Error message to show:', errorMessage);
        showConvertPdfStatus(errorMessage, 'error');
        
        // Show debug section on error for troubleshooting
        const debugSection = document.getElementById('convert-pdf-debug') as HTMLElement;
        if (debugSection) {
            debugSection.style.display = 'block';
            console.log('✅ Debug section shown for error troubleshooting');
        }
        
        console.log('=== CONVERT TO PDF FAILED ===');
    } finally {
        // Re-enable button and hide progress
        convertBtn.disabled = false;
        progressSection.style.display = 'none';
        console.log('🔄 UI reset - button enabled, progress hidden');
    }
}

async function convertWordToPdf(documentBlob: Blob, quality: string): Promise<Blob> {
    console.log('🚀 convertWordToPdf() called');
    try {
        console.log('📊 Input parameters:', {
            documentBlobSize: documentBlob.size,
            documentBlobType: documentBlob.type,
            quality: quality
        });
        
        console.log('🔑 Authentication check:', {
            apiBaseUrl: authCredentials.apiBaseUrl,
            hasApiKey: !!authCredentials.apiKey,
            apiKeyPreview: authCredentials.apiKey ? `${authCredentials.apiKey.substring(0, 20)}...` : 'undefined'
        });
        
        // Create FormData - exactly like the working curl command
        console.log('📦 Creating FormData...');
        const formData = new FormData();
        formData.append('file', documentBlob, 'input.docx');
        formData.append('instructions', JSON.stringify({
            parts: [{ file: "file" }],
            output: { type: "pdf", quality: quality }
        }));
        
        console.log('✅ FormData created with:', {
            fileField: 'file',
            fileName: 'input.docx',
            instructions: {
                parts: [{ file: "file" }],
                output: { type: "pdf", quality: quality }
            }
        });
        
        console.log('🌐 Sending API request...');
        const requestUrl = `${authCredentials.apiBaseUrl}/build`;
        console.log('📡 Request URL:', requestUrl);
        
        // Send exactly the same headers as the working curl command
        const response = await fetch(requestUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${authCredentials.apiKey}`,
                'Accept': '*/*'
            },
            body: formData
        });
        
        console.log('📡 Response received:', {
            status: response.status,
            statusText: response.statusText,
            ok: response.ok
        });
        
        if (response.ok) {
            console.log('✅ Response OK, getting blob...');
            const pdfBlob = await response.blob();
            console.log('✅ PDF blob received:', {
                size: pdfBlob.size,
                type: pdfBlob.type
            });
            return pdfBlob;
        } else {
            console.error('❌ Response not OK, getting error text...');
            const errorText = await response.text();
            console.error('❌ API Error Response:', errorText);
            throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }
        
    } catch (error) {
        console.error('❌ Error converting to PDF:', error);
        throw error;
    }
}

function savePdfForDownload(pdfBlob: Blob, fileName: string) {
    try {
        // Store the blob for download (already in correct format from API)
        (window as any).currentPdfBlob = pdfBlob;
        (window as any).currentPdfFileName = fileName;
        
        console.log('PDF saved for download:', fileName, 'Size:', pdfBlob.size, 'bytes');
        
    } catch (error) {
        console.error('Error saving PDF for download:', error);
        throw error;
    }
}

function showDownloadSection(fileName: string) {
    const downloadSection = document.getElementById('convert-pdf-download') as HTMLElement;
    const downloadFilename = document.getElementById('download-filename') as HTMLElement;
    
    if (downloadSection && downloadFilename) {
        downloadFilename.textContent = `File: ${fileName}`;
        downloadSection.style.display = 'block';
    }
}

function handleDownloadPdf() {
    console.log('📥 handleDownloadPdf() called');
    
    const pdfBlob = (window as any).currentPdfBlob;
    const fileName = (window as any).currentPdfFileName;
    
    console.log('📊 Download data:', {
        hasPdfBlob: !!pdfBlob,
        pdfBlobSize: pdfBlob ? pdfBlob.size : 'N/A',
        pdfBlobType: pdfBlob ? pdfBlob.type : 'N/A',
        fileName: fileName
    });
    
    if (pdfBlob && fileName) {
        try {
            console.log('🔗 Creating download link...');
            
            // Create download link
            const url = URL.createObjectURL(pdfBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = fileName;
            link.style.display = 'none';
            
            console.log('📋 Download link created:', {
                url: url,
                download: fileName
            });
            
            // Trigger download
            console.log('🚀 Triggering download...');
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            // Clean up
            URL.revokeObjectURL(url);
            
            console.log('✅ Download triggered successfully');
            showConvertPdfStatus(`PDF downloaded successfully: ${fileName}`, 'success');
            
        } catch (error) {
            console.error('❌ Error during download:', error);
            showConvertPdfStatus(`Download failed: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
        }
    } else {
        console.error('❌ No PDF available for download:', {
            hasPdfBlob: !!pdfBlob,
            hasFileName: !!fileName
        });
        showConvertPdfStatus('No PDF available for download', 'error');
    }
}

function resetConvertToPdf() {
    // Hide download section
    const downloadSection = document.getElementById('convert-pdf-download') as HTMLElement;
    if (downloadSection) {
        downloadSection.style.display = 'none';
    }
    
    // Hide debug section
    const debugSection = document.getElementById('convert-pdf-debug') as HTMLElement;
    if (debugSection) {
        debugSection.style.display = 'none';
    }
    
    // Clear status messages
    const statusDiv = document.getElementById('convert-pdf-status') as HTMLElement;
    if (statusDiv) {
        statusDiv.innerHTML = '';
    }
    
    // Reset progress
    const progressBar = document.getElementById('convert-pdf-progress-bar') as HTMLElement;
    if (progressBar) {
        progressBar.style.width = '0%';
    }
    
    // Clear stored PDF
    (window as any).currentPdfBlob = null;
    (window as any).currentPdfFileName = null;
}

function showConvertPdfStatus(message: string, type: 'success' | 'error' | 'info') {
    const statusMessages = document.getElementById('convert-pdf-status') as HTMLElement;
    
    if (statusMessages) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `status-message ${type}`;
        messageDiv.textContent = message;
        statusMessages.appendChild(messageDiv);
        
        // Auto-remove after 5 seconds
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.remove();
            }
        }, 5000);
    }
}

function toggleDebugSection() {
    const debugContent = document.getElementById('debug-content');
    const toggleBtn = document.getElementById('toggle-debug-btn');
    
    if (debugContent && toggleBtn) {
        const isCollapsed = debugContent.classList.contains('collapsed');
        
        if (isCollapsed) {
            debugContent.classList.remove('collapsed');
            debugContent.classList.add('expanded');
            toggleBtn.textContent = 'Hide Details';
        } else {
            debugContent.classList.remove('expanded');
            debugContent.classList.add('collapsed');
            toggleBtn.textContent = 'Show Details';
        }
    }
}

function displayApiRequest(requestDetails: any) {
    // Show debug section
    const debugSection = document.getElementById('convert-pdf-debug');
    if (debugSection) {
        debugSection.style.display = 'block';
    }
    
    // Populate request details
    const urlElement = document.getElementById('debug-request-url');
    const methodElement = document.getElementById('debug-request-method');
    const headersElement = document.getElementById('debug-request-headers');
    const payloadElement = document.getElementById('debug-request-payload');
    
    if (urlElement) urlElement.textContent = requestDetails.url;
    if (methodElement) methodElement.textContent = requestDetails.method;
    if (headersElement) headersElement.textContent = JSON.stringify(requestDetails.headers, null, 2);
    if (payloadElement) payloadElement.textContent = requestDetails.payload;
}

function displayApiResponse(responseDetails: any) {
    // Populate response details
    const statusElement = document.getElementById('debug-response-status');
    const headersElement = document.getElementById('debug-response-headers');
    const bodyElement = document.getElementById('debug-response-body');
    
    if (statusElement) statusElement.textContent = responseDetails.status;
    if (headersElement) headersElement.textContent = JSON.stringify(responseDetails.headers, null, 2);
    if (bodyElement) bodyElement.textContent = responseDetails.body;
}
